"""Resumo do mês: recibo da declaração (PGDAS-D) x DAS, por empresa.

Lê as pastas das empresas no servidor (não só o que passou pelo ouvinte) e,
para cada empresa que tem recibo ou DAS na competência, compara:

  FALTA DAS        recibo com débito exigível > 0 e nenhum DAS na pasta
  CONFERIR         DAS sem recibo, DAS com débito zero, DAS menor que o débito
                   ou débito em pasta SEM MOV
  OK - SEM DÉBITO  recibo com débito zero (ou abaixo de R$ 10,00) e sem DAS (esperado)
  OK               recibo com débito > 0 e DAS na pasta

Uso:
  python resumo.py                  competência = mês anterior ao atual
  python resumo.py 09_2026          competência específica

Gera resumos\\recibo_x_das_MM_AAAA.csv (abre no Excel).
"""
import csv
import logging
import os
import re
import sys
from datetime import date
from pathlib import Path

from pypdf import PdfReader

from ouvinte import BASE, carregar_config, pasta_mes_empresa
from regras import mes_anterior, normalizar

V = r"(\d{1,3}(?:\.\d{3})*,\d{2})"
# Tabela "2.1 Apuração no Simples Nacional" do recibo da Receita (vem também dentro do
# PDF "Declaração Mensal_Recibo" da Domínio):
#   MM/AAAA  nº apuração  R$ receita  R$ débito declarado  R$ suspenso  R$ débito exigível
TABELA_RECEITA = re.compile(rf"\d{{2}}/\d{{4}}\s+\d+\s+R\$\s*{V}\s+R\$\s*{V}\s+R\$\s*{V}\s+R\$\s*{V}")
# Reserva: bloco "2.6) Resumo da Declaração" do declaratório; o débito é o último valor.
RESUMO_DECLARACAO = re.compile(r"2\.6\) Resumo da Declara..o(.*?)(?:2\.7\)|$)", re.S)

ORDEM = {"FALTA DAS": 0, "CONFERIR": 1, "OK - SEM DÉBITO": 2, "OK": 3}


def valor(texto):
    return float(texto.replace(".", "").replace(",", "."))


def brl(numero):
    return f"{numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


DAS_CABECALHO = re.compile(r"^\s*Documento de Arrecada..o\s+do Simples Nacional")
DAS_VALOR = re.compile(rf"Valor Total do Documento\s*{V}")
# Extrato do Simples Nacional: no bloco "4) Total Geral da Empresa", a primeira linha de
# valores por tributo é o débito declarado; o último número da linha é o total.
EXTRATO_TOTAL = re.compile(r"4\) Total Geral da Empresa.*?ISS Total\s*\n([\d.,\s]+?)\n", re.S)

DAS_MINIMO = 10.0  # DAS abaixo de R$ 10,00 não é emitido; o valor acumula para o mês seguinte

logging.getLogger("pypdf").setLevel(logging.ERROR)  # PDFs salvos do navegador vêm com xref torto


def ler_pdf(pdf):
    """Identifica o PDF pelo conteúdo (os nomes feitos à mão variam muito:
    PGDAS-DAS.pdf, PGD-DAS.pdf, PGDAS.pdf, RECIBO.pdf, PGDF749_TR.pdf...).
    Retorna ("DAS", valor), ("DECLARAÇÃO", débito) ou (None, None).
    "DECLARAÇÃO" é o recibo do PGDAS-D ou o extrato do Simples Nacional."""
    try:
        paginas = PdfReader(pdf).pages
        primeira = paginas[0].extract_text() or ""
        if DAS_CABECALHO.search(primeira):
            m = DAS_VALOR.search(primeira)
            return "DAS", valor(m.group(1)) if m else None
        topo = normalizar(primeira)
        extrato = "EXTRATO DO SIMPLES NACIONAL" in topo
        if not extrato and "RECIBO" not in normalizar(pdf.name) and "RECIBO" not in topo \
                and "DECLARATORIO" not in topo:
            return None, None
        texto = "\n".join(p.extract_text() or "" for p in paginas)
    except Exception:
        return None, None

    # Recibo de SPED, DCTF, parcelamento etc. não tem esses blocos e é ignorado.
    m = TABELA_RECEITA.search(texto)
    if m:
        return "DECLARAÇÃO", valor(m.group(4))
    if extrato:
        m = EXTRATO_TOTAL.search(texto)
        valores = re.findall(V, m.group(1)) if m else []
        return ("DECLARAÇÃO", valor(valores[-1])) if valores else (None, None)
    m = RESUMO_DECLARACAO.search(texto)
    valores = re.findall(V, m.group(1)) if m else []
    return ("DECLARAÇÃO", valor(valores[-1])) if valores else (None, None)


def analisar(pasta_empresa, ano, mes):
    pasta = pasta_mes_empresa(pasta_empresa, ano, mes)
    if not pasta.is_dir():
        return None
    pdfs = sorted((a for a in pasta.iterdir() if a.is_file() and a.suffix.lower() == ".pdf"),
                  key=lambda a: a.stat().st_mtime)
    das, recibos = [], []
    for pdf in pdfs:
        tipo, quantia = ler_pdf(pdf)
        if tipo == "DAS":
            das.append((pdf, quantia))
        elif tipo == "DECLARAÇÃO":
            recibos.append((pdf, quantia))
    if not das and not recibos:
        return None

    recibo, debito = recibos[-1] if recibos else (None, None)  # o mais recente (retificação / _v2)
    valores_das = [q for _, q in das if q is not None]
    valor_das = max(valores_das) if valores_das else None   # DAS reemitido com juros vale o maior
    sem_mov = "SEM MOV" in pasta.name.upper()

    if not recibo:
        status, obs = "CONFERIR", "DAS sem recibo/extrato da declaração"
    elif 0 < debito < DAS_MINIMO and not das:
        status, obs = "OK - SEM DÉBITO", "débito abaixo de R$ 10,00: acumula para o mês seguinte"
    elif debito > 0 and not das:
        status, obs = "FALTA DAS", ""
    elif debito == 0 and das:
        status, obs = "CONFERIR", "DAS na pasta mas recibo com débito zero"
    elif debito > 0 and sem_mov:
        status, obs = "CONFERIR", "débito declarado em pasta marcada SEM MOV"
    elif debito > 0 and valor_das is not None and valor_das < debito - 0.01:
        status, obs = "CONFERIR", f"DAS de R$ {brl(valor_das)} menor que o débito do recibo"
    elif debito == 0:
        status, obs = "OK - SEM DÉBITO", ""
    else:
        status, obs = "OK", ""

    return {
        "status": status,
        "codigo": int(re.match(r"\d+", pasta_empresa.name).group()),
        "empresa": pasta_empresa.name,
        "debito": debito,
        "valor_das": valor_das,
        "das": " | ".join(p.name for p, _ in das),
        "recibo": recibo.name if recibo else "",
        "observacao": obs,
        "pasta": str(pasta),
    }


def main():
    if len(sys.argv) > 1:
        m = re.fullmatch(r"(\d{1,2})[_/-]?(\d{4})", sys.argv[1])
        if not m:
            sys.exit("Competência inválida. Use MM_AAAA, ex.: 09_2026")
        mes, ano = int(m.group(1)), int(m.group(2))
    else:
        ano, mes = mes_anterior(date.today())
    competencia = f"{mes:02d}_{ano}"
    if sys.stdout is None:  # pythonw (tarefa agendada): o texto do resumo vai para um .txt
        (BASE / "resumos").mkdir(exist_ok=True)
        sys.stdout = open(BASE / "resumos" / f"recibo_x_das_{competencia}.txt", "w", encoding="utf-8")

    raiz = Path(carregar_config()["raiz_empresas"])
    empresas = [Path(e.path) for e in os.scandir(raiz) if e.is_dir() and re.match(r"^\d+\s*-", e.name)]
    linhas = [r for r in (analisar(p, ano, mes) for p in empresas) if r]
    linhas.sort(key=lambda r: (ORDEM[r["status"]], r["codigo"]))

    saida = BASE / "resumos" / f"recibo_x_das_{competencia}.csv"
    saida.parent.mkdir(exist_ok=True)
    with open(saida, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["status", "codigo", "empresa", "debito_exigivel", "valor_das", "observacao", "recibo", "das", "pasta"])
        for r in linhas:
            debito = "" if r["debito"] is None else brl(r["debito"])
            valor_das = "" if r["valor_das"] is None else brl(r["valor_das"])
            w.writerow([r["status"], r["codigo"], r["empresa"], debito, valor_das, r["observacao"],
                        r["recibo"], r["das"], r["pasta"]])

    print(f"Competência {competencia}: {len(linhas)} empresas com recibo ou DAS")
    for status in ORDEM:
        grupo = [r for r in linhas if r["status"] == status]
        print(f"\n{status}: {len(grupo)}")
        if status.startswith("OK"):
            continue
        for r in grupo:
            debito = "-" if r["debito"] is None else f"R$ {brl(r['debito'])}"
            print(f"  {r['empresa']:<28} {debito:>14}  {r['observacao']}")
    print(f"\nPlanilha: {saida}")


if __name__ == "__main__":
    main()
