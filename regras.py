"""Classificação dos arquivos gerados pelas rotinas automáticas da Domínio.

Olha só o nome do arquivo (e a data de geração, quando a competência não vem
no nome) e diz para onde ele vai. Não toca em disco: quem resolve as pastas
reais no servidor é o ouvinte.py.
"""
import os
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime

# Destinos dentro de .RELATÓRIOS (nomes exatamente como estão no servidor)
IMPORTACAO = "IMPORTAÇÃO"
APURACOES = "APURAÇÕES"
ZERAMENTO = "ZERAMENTO"

PASTA_DOC = {"NFS-E": "NFS-e", "NF-E": "NF-e", "NFC-E": "NFC-e", "CT-E E CT-EOS": "CT-e", "CT-E": "CT-e"}


@dataclass
class Classificacao:
    regra: str
    ano: int
    mes: int
    codigo: int | None = None                       # arquivo de empresa -> pasta da empresa
    subpasta: list[str] = field(default_factory=list)  # relatório -> caminho dentro de .RELATÓRIOS
    usuario: str | None = None                      # relatório de usuário -> IMPORTAÇÃO\{n}-{USUÁRIO}

    @property
    def competencia(self):
        return f"{self.mes:02d}_{self.ano}"


class NaoReconhecido(Exception):
    """Arquivo que nenhuma regra soube classificar (vai para _PENDENTES)."""


def normalizar(texto):
    """Maiúsculo e sem acento, para comparar nomes sem depender de Ç, Ã etc."""
    sem = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in sem if not unicodedata.combining(c)).upper().strip()


def mes_anterior(data):
    return (data.year - 1, 12) if data.month == 1 else (data.year, data.month - 1)


def _competencia(mm, aaaa):
    mes, ano = int(mm), int(aaaa)
    if not 1 <= mes <= 12:
        raise NaoReconhecido(f"competência inválida no nome: {mm}{aaaa}")
    return ano, mes


# Arquivos de uma empresa, com competência no nome -> pasta da empresa
REGRAS_EMPRESA = [
    ("DAS", re.compile(r"^DAS - API INTEGRA CONTADOR - EMPRESA (\d+)(?: - .*)? - (\d{2})(\d{4})$")),
    ("RECIBO DAS", re.compile(r"^DECLARACAO MENSAL_RECIBO DAS - .*?EMPRESA (\d+)(?: - .*)? - (\d{2})(\d{4})$")),
    ("LIVRO", re.compile(r"^(?:ENTRADAS|SAIDAS|SERVICOS)_(\d+)_(\d{2})(\d{4})$")),
]

RE_INCONSISTENCIAS = re.compile(r"^RELATORIO DE INCONSISTENCIAS - (.+)$")
RE_NOTAS_LANCADAS = re.compile(r"^.+ - NOTAS LANCADAS$")
RE_IMPORTACAO = re.compile(r"^(NFS-E|NFC-E|NF-E|CT-E E CT-EOS|CT-E)\b.*? - (\d+) - (.+)$")


def classificar(nome_arquivo, gerado_em: datetime, usuarios):
    stem = normalizar(os.path.splitext(nome_arquivo)[0])
    stem = re.sub(r" \[\d{8}-\d{6}\]$", "", stem)  # sufixo que mover_para põe em nome repetido
    ano_ant, mes_ant = mes_anterior(gerado_em)

    for regra, padrao in REGRAS_EMPRESA:
        m = padrao.match(stem)
        if m:
            ano, mes = _competencia(m.group(2), m.group(3))
            return Classificacao(regra, ano, mes, codigo=int(m.group(1)))

    m = RE_INCONSISTENCIAS.match(stem)
    if m:
        return _inconsistencias(m.group(1), ano_ant, mes_ant, usuarios)

    if RE_NOTAS_LANCADAS.match(stem):
        return Classificacao("NOTAS LANÇADAS", ano_ant, mes_ant,
                             subpasta=[IMPORTACAO, "NOTAS LANÇADAS"])

    m = RE_IMPORTACAO.match(stem)
    if m:
        return _importacao(m.group(1), m.group(3), ano_ant, mes_ant)

    raise NaoReconhecido("nome fora dos padrões conhecidos")


def _inconsistencias(descricao, ano, mes, usuarios):
    usuario = next((u for u in usuarios if re.search(rf"\b{re.escape(normalizar(u))}\b", descricao)), None)

    if "ZERAMENTO" in descricao:
        return Classificacao("INCONSISTÊNCIAS ZERAMENTO", ano, mes, subpasta=[ZERAMENTO])
    if "IMPORTACAO" in descricao or "CT-E" in descricao:
        if usuario:
            return Classificacao("INCONSISTÊNCIAS IMPORTAÇÃO", ano, mes,
                                 subpasta=[IMPORTACAO], usuario=usuario)
        if "CT-E" in descricao and "TODOS" in descricao:
            return Classificacao("INCONSISTÊNCIAS CT-e", ano, mes, subpasta=[IMPORTACAO, "CT-e"])
        if "SAIDA" in descricao:
            return Classificacao("INCONSISTÊNCIAS SAÍDAS", ano, mes, subpasta=[IMPORTACAO, "COMERCIO SAÍDA"])
        raise NaoReconhecido("inconsistência de importação sem usuário no nome")
    if any(k in descricao for k in ("APURACAO", "GERADOR DE GUIA", "GUIA PGDAS", "FILIAIS")):
        return Classificacao("INCONSISTÊNCIAS APURAÇÃO", ano, mes, subpasta=[APURACOES])
    raise NaoReconhecido("relatório de inconsistências de rotina desconhecida")


def _importacao(doc, resto, ano, mes):
    # A operação é o último trecho antes de " - GUIA ..."; o nome da empresa vem
    # antes dela e pode conter qualquer palavra, então não entra na decisão.
    antes_guia = resto.split(" - GUIA ")[0]
    trechos = antes_guia.split(" - ")
    operacao = trechos[-1] if len(trechos) > 1 else ""
    if operacao == "EMITIDASRECEBIDAS" and len(trechos) > 2:  # NFS-e Paulistana: "NFS-e Entradas - EmitidasRecebidas"
        operacao = trechos[-2]

    if doc == "NFS-E":
        if "CANCELAD" in operacao:
            categoria = ["SERVIÇO CANCELADO"]
        elif "PRESTAD" in operacao or operacao == "NFS-E SERVICOS":
            categoria = ["SERVIÇO PRESTADO"]
        elif "TOMAD" in operacao or operacao.startswith("NFS-E ENTRADAS"):
            categoria = ["SERVIÇO TOMADO"]
        else:
            raise NaoReconhecido(f"importação NFS-e com operação desconhecida: '{operacao}'")
    else:
        if "CANCELAD" in operacao:
            grupo = "COMERCIO CANCELADO"
        elif "ENTRADA" in operacao:
            grupo = "COMERCIO ENTRADA"
        elif "SAIDA" in operacao:
            grupo = "COMERCIO SAÍDA"
        else:
            raise NaoReconhecido(f"importação {PASTA_DOC[doc]} com operação desconhecida: '{operacao}'")
        categoria = [grupo, PASTA_DOC[doc]]

    return Classificacao(f"IMPORTAÇÃO {PASTA_DOC[doc]}", ano, mes, subpasta=[IMPORTACAO, *categoria])
