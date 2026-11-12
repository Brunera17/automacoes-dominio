"""Separa por empresa os "Relatorio de Inconsistências" da Domínio.

Cada relatório junta várias empresas. Este módulo lê todos os relatórios de
inconsistências em .RELATÓRIOS, agrupa as linhas por empresa e competência e
grava na pasta do mês de cada empresa:

  SIMPLES NACIONAL\\{cód}-NOME\\{AAAA}\\{MM_AAAA}\\Inconsistências Domínio - MM_AAAA.xlsx

A competência de cada relatório é o mês anterior à data de "Emissão" impressa
nele (as rotinas rodam do dia 1 ao 12 sobre o mês anterior). Mensagens
rotineiras ("Nenhuma nota a ser importada" etc.) ficam marcadas na coluna
"Rotineiro" para poderem ser filtradas.

Uso:
  python inconsistencias.py            competência = mês anterior ao atual
  python inconsistencias.py 09_2026    competência específica
O ouvinte chama gerar() sozinho sempre que chega um relatório novo.
"""
import hashlib
import json
import logging
import os
import re
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pypdf import PdfReader

from regras import mes_anterior, normalizar

log = logging.getLogger("ouvinte")
logging.getLogger("pypdf").setLevel(logging.ERROR)

# no .exe (PyInstaller) __file__ fica numa pasta temporária; config/logs/estado ficam ao lado do .exe
BASE = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
ESTADO = BASE / "estado"
CACHE = ESTADO / "inconsistencias_cache.json"      # relatório já lido -> linhas
GERADOS = ESTADO / "inconsistencias_gerados.json"  # competência -> empresa -> hash do conteúdo gravado

# Mensagens que aparecem para quase todas as empresas e raramente pedem ação.
ROTINEIRAS = [
    "NENHUMA NOTA A SER IMPORTADA",
    "INFORMAR A DATA DE ENTRADA NO MOMENTO DA IMPORTACAO",
    "NAO HA REGISTROS VALIDOS PARA GRAVAR",
    "DADOS GRAVADOS COM SUCESSO",
    "SEM DADOS PARA EMITIR",
]

# Colunas do relatório (posição x do texto na página, em pontos)
X_MENSAGEM = 110   # antes disso: "código Tipo"
X_DIREITA = 400   # depois disso: Imposto / Período (Imposto começa em ~412)
RE_REGISTRO = re.compile(r"^(?:(\d+)\s+)?(Erro|Aviso)\b\s*(.*)$")
RE_DATA = re.compile(r"^\d{2}/\d{2}/\d{4}$")
IGNORAR = ("Sistema licenciado", "Página", "Emissão", "Hora")


# ---------------------------------------------------------------- leitura

def _itens(pagina):
    itens = []

    def visitor(texto, cm, tm, _fd, _fs):
        texto = " ".join(texto.split())
        if texto:
            itens.append((cm[5] + tm[5], cm[4] + tm[4], texto))

    pagina.extract_text(visitor_text=visitor)
    return sorted(itens, key=lambda i: (-i[0], i[1]))


def _valor_ao_lado(itens, rotulo):
    for y, _x, t in itens:
        if t == rotulo:
            vizinhos = [(x, t2) for y2, x, t2 in itens if abs(y2 - y) < 1.5 and t2 != rotulo]
            return min(vizinhos)[1] if vizinhos else ""
    return ""


def ler_relatorio(pdf):
    """Lê um relatório de inconsistências. Retorna (data de emissão, linhas).

    O texto da coluna Mensagem não vem alinhado com a coluna Empresa no PDF; por isso
    a associação é feita pela posição: cada linha de mensagem pertence ao registro
    (código + tipo) imediatamente acima dela. A primeira linha de mensagem pode ficar
    até ~1pt acima do código, daí a tolerância de 2.5pt."""
    registros, emissao, rotina = [], None, ""
    for pagina in PdfReader(pdf).pages:
        itens = _itens(pagina)
        emissao = emissao or _valor_ao_lado(itens, "Emissão:")
        cabecalhos = [y for y, _x, t in itens if t.startswith("Empresa Tipo")]
        if not cabecalhos:
            continue

        # Uma página pode ter várias seções "Rotina: ..." empilhadas; cada registro
        # pertence à última rotina impressa acima dele (ou à da página anterior).
        rotinas = []
        for y, _x, t in itens:
            if t == "Rotina:":
                valor = [(x2, t2) for y2, x2, t2 in itens if abs(y2 - y) < 1.5 and t2 != "Rotina:"]
                nome = min(valor)[1] if valor else ""
                rotinas.append((y, f"Rotina {nome}" if nome.isdigit() else nome))  # às vezes vem só o código
        linhas_de_secao = {y for y, _ in rotinas} | set(cabecalhos)

        corpo = [(y, x, t) for y, x, t in itens
                 if y < max(cabecalhos) - 1 and not t.startswith(IGNORAR)
                 and not any(abs(y - ys) < 1.5 for ys in linhas_de_secao)]
        da_pagina = []  # (y, registro)
        for y, x, t in corpo:
            m = RE_REGISTRO.match(t) if x < X_MENSAGEM else None
            if not m:
                continue
            codigo = int(m.group(1)) if m.group(1) else (registros[-1]["codigo"] if registros else None)
            if codigo is None:
                continue
            acima = [(yr, nome) for yr, nome in rotinas if yr > y]
            if acima:
                rotina = min(acima)[1]
            texto, periodo = m.group(3), ""
            fim = re.search(r"\s(\d{2}/\d{2}/\d{4})$", texto)  # período às vezes vem colado na mensagem
            if fim:
                texto, periodo = texto[:fim.start()], fim.group(1)
            reg = {"codigo": codigo, "rotina": rotina, "tipo": m.group(2),
                   "linhas": [texto] if texto else [], "imposto": "", "periodo": periodo}
            registros.append(reg)
            da_pagina.append((y, reg))

        for y, x, t in corpo:
            if x < X_MENSAGEM:
                continue
            acima = [(yc, r) for yc, r in da_pagina if yc >= y - 2.5]
            if acima:
                reg = min(acima, key=lambda a: a[0])[1]
            elif registros and registros[-1] not in [r for _, r in da_pagina]:
                reg = registros[-1]          # mensagem que continua da página anterior
            elif da_pagina:
                reg = da_pagina[0][1]
            else:
                continue
            if x >= X_DIREITA:
                reg["periodo" if RE_DATA.match(t) else "imposto"] = t
            else:
                reg["linhas"].append(t)

    linhas = []
    for r in registros:
        mensagem = " ".join(r["linhas"]).strip()
        linhas.append({
            "codigo": r["codigo"], "rotina": r["rotina"], "tipo": r["tipo"], "mensagem": mensagem,
            "imposto": r["imposto"], "periodo": r["periodo"],
            "rotineiro": any(p in normalizar(mensagem) for p in ROTINEIRAS),
        })
    return emissao, linhas


def _competencia_da_emissao(emissao):
    data = datetime.strptime(emissao, "%d/%m/%Y")
    ano, mes = mes_anterior(data)
    return f"{mes:02d}_{ano}"


def _carregar(caminho):
    try:
        return json.loads(caminho.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return {}


def _salvar(caminho, dados):
    ESTADO.mkdir(exist_ok=True)
    tmp = caminho.with_suffix(".tmp")
    tmp.write_text(json.dumps(dados, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, caminho)


def ler_todos(raiz_relatorios):
    """Lê (com cache) todos os relatórios de inconsistências em .RELATÓRIOS."""
    if not Path(raiz_relatorios).is_dir():
        raise OSError(f"pasta de relatórios inacessível: {raiz_relatorios}")
    cache, mudou = _carregar(CACHE), False
    vistos = set()
    for pdf in Path(raiz_relatorios).rglob("Relatorio de Inconsist*.pdf"):
        chave = str(pdf)
        vistos.add(chave)
        st = pdf.stat()
        item = cache.get(chave)
        if item and item["mtime"] == st.st_mtime and item["tamanho"] == st.st_size:
            continue
        try:
            emissao, linhas = ler_relatorio(pdf)
            competencia = _competencia_da_emissao(emissao)
        except Exception as e:
            log.warning("Não foi possível ler %s: %s", pdf.name, e)
            emissao, linhas, competencia = "", [], ""
        cache[chave] = {"mtime": st.st_mtime, "tamanho": st.st_size, "emissao": emissao,
                        "competencia": competencia, "linhas": linhas}
        mudou = True
    for chave in set(cache) - vistos:
        if not Path(chave).exists():  # só esquece relatório que de fato sumiu
            del cache[chave]
            mudou = True
    if mudou:
        _salvar(CACHE, cache)
    return cache


# ---------------------------------------------------------------- gravação

COLUNAS = [("Rotina", 34), ("Relatório", 44), ("Emissão", 11), ("Tipo", 7), ("Mensagem", 90),
           ("Imposto", 10), ("Período", 11), ("Rotineiro", 10)]


def _gravar_xlsx(destino, competencia, empresa, linhas):
    wb = Workbook()
    ws = wb.active
    ws.title = competencia
    ws.append([c for c, _ in COLUNAS])
    for i, (_, largura) in enumerate(COLUNAS, start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = largura
        ws.cell(1, i).font = Font(bold=True, color="FFFFFF")
        ws.cell(1, i).fill = PatternFill("solid", fgColor="44546A")
    cinza = Font(color="808080")
    for l in linhas:
        ws.append([l["rotina"], l["relatorio"], l["emissao"], l["tipo"], l["mensagem"],
                   l["imposto"], l["periodo"], "Sim" if l["rotineiro"] else "Não"])
        linha = ws.max_row
        ws.cell(linha, 5).alignment = Alignment(wrap_text=True, vertical="top")
        if l["rotineiro"]:
            for c in range(1, len(COLUNAS) + 1):
                ws.cell(linha, c).font = cinza
        elif l["tipo"] == "Erro":
            ws.cell(linha, 4).font = Font(bold=True, color="C00000")
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    ws.oddFooter.left.text = empresa
    destino.parent.mkdir(parents=True, exist_ok=True)
    tmp = destino.with_name(destino.stem + ".tmp.xlsx")
    wb.save(tmp)
    os.replace(tmp, destino)


def gerar(competencia, cfg, indice, simulacao=False):
    """Grava/atualiza o arquivo de inconsistências de cada empresa na competência.
    Só regrava a empresa cujo conteúdo mudou. Retorna (gravados, pendências)."""
    from ouvinte import Pendente, pasta_mes_empresa  # import tardio: ouvinte importa este módulo

    mes, ano = int(competencia[:2]), int(competencia[3:])
    cache = ler_todos(cfg["raiz_relatorios"])

    por_empresa = {}
    for caminho, item in cache.items():
        if item["competencia"] != competencia:
            continue
        for l in item["linhas"]:
            linha = dict(l, relatorio=Path(caminho).stem.replace("Relatorio de Inconsistências - ", ""),
                         emissao=item["emissao"])
            por_empresa.setdefault(l["codigo"], []).append(linha)

    gerados = _carregar(GERADOS)
    da_competencia = gerados.setdefault(competencia, {})
    gravados, pendencias = [], []
    indice.atualizar()
    for codigo, linhas in sorted(por_empresa.items()):
        # sem duplicar (relatório _v2 idêntico) e com o que exige ação primeiro
        unicas = {(l["rotina"], l["tipo"], l["mensagem"], l["imposto"], l["periodo"]): l for l in linhas}
        linhas = sorted(unicas.values(), key=lambda l: (l["rotineiro"], l["tipo"] != "Erro", l["rotina"]))
        assinatura = hashlib.sha256(json.dumps(linhas, sort_keys=True, ensure_ascii=False).encode()).hexdigest()

        try:
            pasta_empresa = indice.pasta(codigo)
        except Pendente as e:
            pendencias.append(str(e))
            continue
        destino = pasta_mes_empresa(pasta_empresa, ano, mes) / f"Inconsistências Domínio - {competencia}.xlsx"
        if da_competencia.get(str(codigo)) == assinatura and destino.exists():
            continue
        if simulacao:
            gravados.append(destino)
            continue
        try:
            _gravar_xlsx(destino, competencia, pasta_empresa.name, linhas)
        except OSError as e:  # ex.: planilha aberta no Excel; tenta de novo no próximo relatório
            pendencias.append(f"{destino.name} da empresa {codigo} não gravado: {e}")
            continue
        da_competencia[str(codigo)] = assinatura
        gravados.append(destino)

    if not simulacao:
        _salvar(GERADOS, gerados)
    return gravados, pendencias


def main():
    from ouvinte import IndiceEmpresas, carregar_config, configurar_log

    if len(sys.argv) > 1:
        m = re.fullmatch(r"(\d{1,2})[_/-]?(\d{4})", sys.argv[1])
        if not m:
            sys.exit("Competência inválida. Use MM_AAAA, ex.: 09_2026")
        competencia = f"{int(m.group(1)):02d}_{m.group(2)}"
    else:
        ano, mes = mes_anterior(date.today())
        competencia = f"{mes:02d}_{ano}"

    configurar_log()
    cfg = carregar_config()
    gravados, pendencias = gerar(competencia, cfg, IndiceEmpresas(cfg["raiz_empresas"]), cfg.get("simulacao", False))
    log.info("Inconsistências %s: %d arquivo(s) de empresa gravado(s)", competencia, len(gravados))
    for p in sorted(set(pendencias)):
        log.warning("Inconsistências %s: %s", competencia, p)


if __name__ == "__main__":
    main()
