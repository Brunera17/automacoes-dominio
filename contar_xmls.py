"""Conta os XMLs fiscais de cada empresa do servidor, por tipo, direção e mês.

Percorre as pastas no padrão {cód}-APELIDO\\AAAA\\MM_AAAA\\... e gera uma
planilha com o volume real de cada carteira, para dimensionar a importação.

Como classifica cada XML:
- tipo: pelo modelo da chave de acesso no nome do arquivo (55 NF-e, 65 NFC-e,
  57 CT-e, 58 MDF-e, 59 CF-e SAT). Sem chave no nome, lê o começo do arquivo
  e decide pela tag raiz (NFS-e, eventos etc.).
- direção: pelo nome das pastas (ENTRADA/TOMADO x SAÍDA/PRESTADO). Se a pasta
  não diz, compara o CNPJ emitente da chave com o CNPJ da empresa, deduzido
  como o emitente mais frequente nas pastas de saída.
- mês: pela pasta MM_AAAA (competência); se não houver, pelo AAMM da chave.
- duplicatas: a mesma chave contada uma vez só por empresa (XML solto + cópia
  dentro de .zip, ou o mesmo XML em duas pastas).

Não altera nada no servidor: só lê.

Uso:
    python contar_xmls.py                       # todas as empresas, todos os meses
    python contar_xmls.py --desde 2025-01       # só competências a partir de jan/2025
    python contar_xmls.py --empresas 107 559    # só esses códigos
    python contar_xmls.py --sem-zip             # não abre os .zip
"""
import argparse
import json
import os
import re
import sys
import time
import unicodedata
import zipfile
from collections import Counter, defaultdict
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

PASTA = os.path.dirname(os.path.abspath(__file__))

RE_EMPRESA = re.compile(r"^(\d+)-(.+)$")
RE_COMPETENCIA = re.compile(r"^(\d{2})_(\d{4})$")
RE_CHAVE = re.compile(r"(?<!\d)(\d{44})(?!\d)")
# Uma nota dentro de um XML de NFS-e (lotes da prefeitura trazem várias por arquivo)
RE_NOTA_NFSE = re.compile(rb"<(?:\w+:)?(?:nota|CompNfse|Nfse|NFSe|tcNfse)[\s>]")
RE_CNPJ_PRESTADOR = re.compile(rb"<(?:\w+:)?(?:cpfCnpjPrestador|PrestadorServico|Prestador|prestador)\b.*?(\d{14})", re.S)

MODELOS = {"55": "NF-e", "65": "NFC-e", "57": "CT-e", "58": "MDF-e", "59": "CF-e SAT", "67": "CT-e OS"}

# Tag raiz (ou trecho dela) -> tipo, para XMLs sem chave no nome
TAGS_RAIZ = [
    ("procEvento", "Evento"),
    ("envEvento", "Evento"),
    ("retEvento", "Evento"),
    ("nfeProc", "NF-e"),
    ("cteProc", "CT-e"),
    ("mdfeProc", "MDF-e"),
    ("CFe", "CF-e SAT"),
    ("listaNota", "NFS-e"),
    ("CompNfse", "NFS-e"),
    ("ConsultarNfse", "NFS-e"),
    ("Nfse", "NFS-e"),
    ("NFSe", "NFS-e"),
    ("nfse", "NFS-e"),
    ("RPS", "NFS-e"),
    ("Rps", "NFS-e"),
    ("NFe", "NF-e"),
]

PALAVRAS_ENTRADA = ("ENTRADA", "TOMAD", "COMPRA", "RECEBID")
PALAVRAS_SAIDA = ("SAIDA", "PRESTAD", "EMITID", "VENDA")


def normalizar(texto):
    """Maiúsculo e sem acento, para comparar nomes sem depender de Ç, Ã etc."""
    sem_acento = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    return sem_acento.upper()


def tipo_pelo_conteudo(inicio):
    """Decide o tipo pelos primeiros bytes do XML (tag raiz e vizinhança)."""
    texto = inicio.decode("utf-8", "ignore")
    m = re.search(r"<(?!\?)(?:\w+:)?(\w+)", texto)
    raiz = m.group(1) if m else ""
    for trecho, tipo in TAGS_RAIZ:
        if trecho in raiz:
            return tipo
    for trecho, tipo in TAGS_RAIZ:  # raiz genérica (ex.: <root>): procura no começo do arquivo
        if f"<{trecho}" in texto or f":{trecho}" in texto:
            return tipo
    return "Outro XML"


def direcao_pelo_caminho(partes):
    for parte in reversed(partes):
        p = normalizar(parte)
        if any(w in p for w in PALAVRAS_ENTRADA):
            return "Entrada"
        if any(w in p for w in PALAVRAS_SAIDA):
            return "Saída"
    return None


def competencia_pelo_caminho(partes):
    for parte in partes:
        m = RE_COMPETENCIA.match(parte)
        if m and 1 <= int(m.group(1)) <= 12:
            return f"{m.group(2)}-{m.group(1)}"
    return None


class Registro:
    __slots__ = ("chave", "tipo", "direcao", "competencia", "cancelada", "origem", "id_dedup", "notas", "emitente")

    def __init__(self, chave, tipo, direcao, competencia, cancelada, origem, id_dedup, notas=1, emitente=None):
        self.chave = chave
        self.tipo = tipo
        self.direcao = direcao
        self.competencia = competencia
        self.cancelada = cancelada
        self.origem = origem
        self.id_dedup = id_dedup
        self.notas = notas          # documentos dentro do arquivo (lote de NFS-e pode ter vários)
        self.emitente = emitente    # CNPJ emitente/prestador, para deduzir a direção


def registro_de(nome, partes, ler, origem):
    """Monta o registro de um XML a partir do nome, das pastas e (se preciso) do conteúdo.

    ler(n) devolve os n primeiros bytes do arquivo (None = arquivo inteiro).
    """
    ler_inicio = lambda: ler(2048)
    nome_norm = normalizar(nome)
    m = RE_CHAVE.search(nome)
    chave = m.group(1) if m else None
    eh_evento = "EVENTO" in nome_norm or nome_norm.startswith("ID1101") or "-CANC" in nome_norm

    if chave and not eh_evento:
        tipo = MODELOS.get(chave[20:22]) or tipo_pelo_conteudo(ler_inicio())
    elif eh_evento:
        tipo = "Evento"
    else:
        tipo = tipo_pelo_conteudo(ler_inicio())

    competencia = competencia_pelo_caminho(partes)
    if not competencia and chave:
        aa, mm = chave[2:4], chave[4:6]
        if 1 <= int(mm) <= 12:
            competencia = f"20{aa}-{mm}"

    caminho_norm = normalizar("/".join(partes))
    cancelada = "CANCEL" in caminho_norm or "CANCEL" in nome_norm
    notas, emitente = 1, chave[6:20] if chave else None
    if tipo == "NFS-e" and not chave:
        conteudo = ler(None)
        notas = max(1, len(RE_NOTA_NFSE.findall(conteudo)))
        m = RE_CNPJ_PRESTADOR.search(conteudo)
        emitente = m.group(1).decode() if m else None

    # Eventos repetem a chave da nota, então não entram na deduplicação por chave
    id_dedup = chave if chave and tipo != "Evento" else (tipo, nome_norm)
    return Registro(chave, tipo, direcao_pelo_caminho(partes), competencia, cancelada, origem, id_dedup, notas, emitente)


def varrer_empresa(raiz_empresa, abrir_zip, erros):
    """Lista todos os XMLs da empresa (soltos e dentro de .zip)."""
    registros = []
    pilha = [(raiz_empresa, [])]
    while pilha:
        pasta, partes = pilha.pop()
        try:
            entradas = list(os.scandir(pasta))
        except OSError as e:
            erros.append(f"{pasta}: {e}")
            continue
        for ent in entradas:
            try:
                if ent.is_dir(follow_symlinks=False):
                    pilha.append((ent.path, partes + [ent.name]))
                    continue
            except OSError:
                continue
            nome_low = ent.name.lower()
            if nome_low.endswith(".xml"):
                def ler(n, caminho=ent.path):
                    try:
                        with open(caminho, "rb") as f:
                            return f.read(n or 20_000_000)
                    except OSError:
                        return b""
                registros.append(registro_de(ent.name, partes, ler, "solto"))
            elif abrir_zip and nome_low.endswith(".zip"):
                registros.extend(xmls_do_zip(ent.path, partes, erros))
    return registros


def xmls_do_zip(caminho, partes, erros):
    registros = []
    try:
        with zipfile.ZipFile(caminho) as z:
            for info in z.infolist():
                if info.is_dir() or not info.filename.lower().endswith(".xml"):
                    continue
                internas = [p for p in info.filename.replace("\\", "/").split("/") if p]
                nome = internas[-1]

                def ler(n, info=info, z=z):
                    try:
                        with z.open(info) as f:
                            return f.read(n or 20_000_000)
                    except Exception:
                        return b""
                registros.append(registro_de(nome, partes + internas[:-1], ler, "zip"))
    except (zipfile.BadZipFile, OSError, RuntimeError) as e:
        erros.append(f"{caminho}: {e}")
    return registros


def classificar_direcao(registros):
    """Preenche a direção que a pasta não disse, usando o CNPJ da própria empresa."""
    emitentes_saida = Counter(r.emitente for r in registros if r.emitente and r.direcao == "Saída")
    emitentes = emitentes_saida or Counter(r.emitente for r in registros if r.emitente and r.tipo in ("NFC-e", "CF-e SAT", "NFS-e"))
    cnpj_empresa = emitentes.most_common(1)[0][0] if emitentes else None
    for r in registros:
        if r.direcao:
            continue
        if r.tipo in ("NFC-e", "CF-e SAT"):
            r.direcao = "Saída"  # consumidor final: sempre emitida pela empresa
        elif r.emitente and cnpj_empresa:
            r.direcao = "Saída" if r.emitente == cnpj_empresa else "Entrada"
        else:
            r.direcao = "Indefinida"
    return cnpj_empresa


def deduplicar(registros):
    vistos = {}
    for r in registros:
        atual = vistos.get(r.id_dedup)
        # Prefere o XML solto à cópia do zip, e a versão marcada como cancelada
        if atual is None or (atual.origem == "zip" and r.origem == "solto") or (r.cancelada and not atual.cancelada):
            vistos[r.id_dedup] = r
    return list(vistos.values()), len(registros) - len(vistos)


def listar_empresas(raiz, codigos):
    empresas, sem_codigo = [], []
    for ent in sorted(os.scandir(raiz), key=lambda e: e.name):
        if not ent.is_dir() or ent.name.startswith((".", "_")):
            continue
        m = RE_EMPRESA.match(ent.name)
        if not m:
            sem_codigo.append(ent.name)
            continue
        codigo = int(m.group(1))
        if codigos and codigo not in codigos:
            continue
        empresas.append((codigo, m.group(2).strip(), ent.path))
    empresas.sort(key=lambda e: e[0])
    return empresas, sem_codigo


# ---------------------------------------------------------------- planilha

CABECALHO = Font(bold=True, color="FFFFFF")
FUNDO = PatternFill("solid", fgColor="305496")


def escrever_aba(ws, cabecalho, linhas, larguras=None):
    ws.append(cabecalho)
    for c in ws[1]:
        c.font, c.fill = CABECALHO, FUNDO
    for linha in linhas:
        ws.append(linha)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, titulo in enumerate(cabecalho, 1):
        largura = (larguras or {}).get(titulo, max(12, len(titulo) + 2))
        ws.column_dimensions[get_column_letter(i)].width = largura


def gerar_planilha(destino, resultados, tipos, meses, sem_codigo, erros, parametros):
    wb = Workbook()

    # Resumo por empresa
    ws = wb.active
    ws.title = "Por empresa"
    cab = ["Código", "Apelido", "CNPJ deduzido", "Arquivos XML", "Total documentos", "Meses com XML", "Média/mês", "Pico/mês",
           "Mês do pico", "Entradas", "Saídas", "Dir. indefinida", "Canceladas", "Duplicatas ignoradas", "Tempo varredura (s)"]
    cab += [f"Tipo: {t}" for t in tipos]
    linhas = []
    for e in resultados:
        por_mes = Counter()
        for (comp, _t, _d), q in e["contagem"].items():
            por_mes[comp] += q
        total = sum(por_mes.values())
        meses_com = len([m for m in por_mes if m != "sem mês"])
        pico_mes, pico = max(por_mes.items(), key=lambda kv: kv[1]) if por_mes else ("", 0)
        dirs = Counter()
        tipos_emp = Counter()
        for (_c, t, d), q in e["contagem"].items():
            dirs[d] += q
            tipos_emp[t] += q
        linhas.append([e["codigo"], e["apelido"], e["cnpj"] or "", e["arquivos"], total, meses_com,
                       round(total / meses_com, 1) if meses_com else 0, pico, pico_mes,
                       dirs["Entrada"], dirs["Saída"], dirs["Indefinida"], e["canceladas"], e["duplicatas"],
                       round(e["tempo"], 1)] + [tipos_emp[t] for t in tipos])
    linhas.sort(key=lambda l: -l[4])
    escrever_aba(ws, cab, linhas, {"Apelido": 28, "CNPJ deduzido": 18})

    # Matriz empresa x mês
    ws = wb.create_sheet("Empresa x mês")
    linhas = []
    for e in resultados:
        por_mes = Counter()
        for (comp, _t, _d), q in e["contagem"].items():
            por_mes[comp] += q
        linhas.append([e["codigo"], e["apelido"]] + [por_mes[m] for m in meses] + [sum(por_mes.values())])
    escrever_aba(ws, ["Código", "Apelido"] + meses + ["Total"], linhas, {"Apelido": 28})

    # Total da carteira por mês e tipo
    ws = wb.create_sheet("Carteira por mês")
    total = defaultdict(Counter)
    for e in resultados:
        for (comp, t, _d), q in e["contagem"].items():
            total[comp][t] += q
    linhas = [[m] + [total[m][t] for t in tipos] + [sum(total[m].values()), len([e for e in resultados if any(k[0] == m for k in e["contagem"])])]
              for m in meses]
    escrever_aba(ws, ["Competência"] + tipos + ["Total", "Empresas com XML"], linhas)

    # Detalhe (formato longo, bom para tabela dinâmica)
    ws = wb.create_sheet("Detalhe")
    linhas = []
    for e in resultados:
        for (comp, t, d), q in sorted(e["contagem"].items()):
            linhas.append([e["codigo"], e["apelido"], comp, t, d, q])
    escrever_aba(ws, ["Código", "Apelido", "Competência", "Tipo", "Direção", "Quantidade"], linhas, {"Apelido": 28})

    # Informações da execução
    ws = wb.create_sheet("Execução")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 100
    for k, v in parametros.items():
        ws.append([k, str(v)])
    ws.append([])
    ws.append(["Pastas sem código (ignoradas)", len(sem_codigo)])
    for nome in sem_codigo:
        ws.append(["", nome])
    ws.append([])
    ws.append(["Erros de leitura", len(erros)])
    for err in erros[:2000]:
        ws.append(["", err])

    wb.save(destino)


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    with open(os.path.join(PASTA, "config.json"), encoding="utf-8") as f:
        config = json.load(f)

    ap = argparse.ArgumentParser(description="Conta XMLs por empresa, tipo e mês no servidor.")
    ap.add_argument("--raiz", default=config["raiz_empresas"], help="pasta com as empresas (padrão: config.json)")
    ap.add_argument("--desde", help="primeira competência, AAAA-MM")
    ap.add_argument("--ate", help="última competência, AAAA-MM")
    ap.add_argument("--empresas", nargs="*", type=int, help="códigos das empresas (padrão: todas)")
    ap.add_argument("--sem-zip", action="store_true", help="não abre os .zip")
    ap.add_argument("--saida", help="caminho do .xlsx (padrão: .Ouvinte\\contagens\\XMLs_AAAAMMDD_HHMM.xlsx)")
    args = ap.parse_args()

    empresas, sem_codigo = listar_empresas(args.raiz, set(args.empresas or []))
    if not empresas:
        sys.exit(f"Nenhuma empresa no padrão código-apelido em {args.raiz}")

    print(f"{len(empresas)} empresas em {args.raiz} ({len(sem_codigo)} pastas sem código ignoradas)")
    inicio_total = time.perf_counter()
    resultados, erros = [], []
    tipos_vistos, meses_vistos = Counter(), set()

    for i, (codigo, apelido, caminho) in enumerate(empresas, 1):
        t0 = time.perf_counter()
        registros = varrer_empresa(caminho, not args.sem_zip, erros)
        cnpj = classificar_direcao(registros)
        registros, duplicatas = deduplicar(registros)

        contagem = Counter()
        canceladas = arquivos = 0
        for r in registros:
            comp = r.competencia or "sem mês"
            if comp != "sem mês" and ((args.desde and comp < args.desde) or (args.ate and comp > args.ate)):
                continue
            contagem[(comp, r.tipo, r.direcao)] += r.notas
            arquivos += 1
            canceladas += r.cancelada
            tipos_vistos[r.tipo] += r.notas
            meses_vistos.add(comp)

        tempo = time.perf_counter() - t0
        resultados.append({"codigo": codigo, "apelido": apelido, "cnpj": cnpj, "contagem": contagem,
                           "canceladas": canceladas, "duplicatas": duplicatas, "tempo": tempo, "arquivos": arquivos})
        print(f"[{i}/{len(empresas)}] {codigo}-{apelido}: {sum(contagem.values())} documentos em {arquivos} XMLs ({tempo:.1f}s)", flush=True)

    duracao = time.perf_counter() - inicio_total
    tipos = [t for t, _ in tipos_vistos.most_common()]
    meses = sorted(m for m in meses_vistos if m != "sem mês") + (["sem mês"] if "sem mês" in meses_vistos else [])

    pasta_saida = os.path.join(PASTA, "contagens")
    os.makedirs(pasta_saida, exist_ok=True)
    destino = args.saida or os.path.join(pasta_saida, f"XMLs_{datetime.now():%Y%m%d_%H%M}.xlsx")
    parametros = {
        "Gerado em": f"{datetime.now():%d/%m/%Y %H:%M}",
        "Raiz": args.raiz,
        "Competências": f"{args.desde or 'início'} a {args.ate or 'hoje'}",
        "Abriu .zip": not args.sem_zip,
        "Empresas varridas": len(empresas),
        "Arquivos XML": sum(e["arquivos"] for e in resultados),
        "Total de documentos": sum(tipos_vistos.values()),
        "Obs.": "Documentos = notas/eventos; um XML de lote de NFS-e da prefeitura pode ter várias notas.",
        "Duração da varredura (s)": round(duracao, 1),
    }
    gerar_planilha(destino, resultados, tipos, meses, sem_codigo, erros, parametros)

    print(f"\nTotal: {sum(tipos_vistos.values())} documentos em {sum(e['arquivos'] for e in resultados)} XMLs, {duracao:.0f}s")
    for t, q in tipos_vistos.most_common():
        print(f"  {t}: {q}")
    if erros:
        print(f"{len(erros)} erros de leitura (ver aba Execução)")
    print(f"Planilha: {destino}")


if __name__ == "__main__":
    main()
