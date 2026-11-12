"""Ouvinte do hub de automações da Domínio.

Fica olhando a pasta de entrada (onde as rotinas automáticas salvam os PDFs),
classifica cada arquivo pelo nome (regras.py) e copia para o servidor:

  - arquivo de empresa  -> SIMPLES NACIONAL\\{cód}-NOME\\{AAAA}\\{MM_AAAA}\\
  - relatório de rotina -> SIMPLES NACIONAL\\.RELATÓRIOS\\...\\{MM_AAAA}\\

Nenhum arquivo some: ou vai para o servidor (cópia conferida por hash e o
original arquivado em _PROCESSADOS), ou fica em _PENDENTES / _ERROS com o
motivo registrado em logs\\movimentos_AAAA-MM.csv.

Uso:
  python ouvinte.py                 roda em loop (modo definido no config.json)
  python ouvinte.py --uma-vez       roda um ciclo e sai
  python ouvinte.py --simular       não copia nem move nada, só registra o que faria
  python ouvinte.py --real          ignora "simulacao" do config e move de verdade
"""
import argparse
import csv
import ctypes
import hashlib
import json
import logging
import os
import re
import shutil
import sys
import time
from datetime import datetime, timedelta
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path

import inconsistencias
from regras import NaoReconhecido, classificar

# no .exe (PyInstaller) __file__ fica numa pasta temporária; config/logs/estado ficam ao lado do .exe
BASE = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
LOGS = BASE / "logs"
IGNORAR = re.compile(r"^(~\$|desktop\.ini$|thumbs\.db$)|\.(tmp|part|crdownload)$", re.I)

log = logging.getLogger("ouvinte")


class Pendente(Exception):
    """O arquivo foi reconhecido, mas o destino não pôde ser determinado."""


def carregar_config():
    with open(BASE / "config.json", encoding="utf-8") as f:
        return json.load(f)


def configurar_log():
    LOGS.mkdir(exist_ok=True)
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s", "%d/%m/%Y %H:%M:%S")
    arquivo = TimedRotatingFileHandler(LOGS / "ouvinte.log", when="midnight", backupCount=60, encoding="utf-8")
    arquivo.setFormatter(fmt)
    log.addHandler(arquivo)
    if sys.stdout:  # pythonw (tarefa agendada) roda sem console
        tela = logging.StreamHandler(sys.stdout)
        tela.setFormatter(fmt)
        log.addHandler(tela)
    log.setLevel(logging.INFO)


def registrar(status, origem, classif=None, destino="", motivo=""):
    """Uma linha por evento no CSV do mês (abre direto no Excel)."""
    caminho = LOGS / f"movimentos_{datetime.now():%Y-%m}.csv"
    novo = not caminho.exists()
    with open(caminho, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        if novo:
            w.writerow(["data_hora", "status", "arquivo", "regra", "empresa", "competencia", "destino", "motivo"])
        w.writerow([
            f"{datetime.now():%d/%m/%Y %H:%M:%S}", status, Path(origem).name,
            classif.regra if classif else "", (classif.codigo or "") if classif else "",
            classif.competencia if classif else "", destino, motivo,
        ])
    nivel = logging.WARNING if status in ("PENDENTE", "ERRO", "AVISO") else logging.INFO
    log.log(nivel, "%s | %s%s%s", status, Path(origem).name,
            f" -> {destino}" if destino else "", f" ({motivo})" if motivo else "")


# ---------------------------------------------------------------- servidor

class IndiceEmpresas:
    """Mapa código -> pasta da empresa, montado lendo SIMPLES NACIONAL."""

    def __init__(self, raiz, validade_seg=300):
        self.raiz = Path(raiz)
        self.validade = validade_seg
        self.lido_em = 0
        self.mapa = {}

    def atualizar(self, forcar=False):
        if not forcar and time.time() - self.lido_em < self.validade:
            return
        mapa = {}
        for item in os.scandir(self.raiz):
            m = re.match(r"^(\d+)\s*-", item.name)
            if m and item.is_dir():
                mapa.setdefault(int(m.group(1)), []).append(Path(item.path))
        self.mapa, self.lido_em = mapa, time.time()

    def pasta(self, codigo):
        pastas = self.mapa.get(codigo, [])
        if not pastas:
            raise Pendente(f"empresa {codigo} sem pasta '{codigo}-...' em {self.raiz.name}")
        if len(pastas) > 1:
            raise Pendente(f"código {codigo} duplicado no servidor: " + " | ".join(p.name for p in pastas))
        return pastas[0]


def pasta_mes_empresa(pasta_empresa, ano, mes):
    """Reaproveita a pasta do mês que já existir (09_2026, 09_2026 SEM MOV, 09-2026,
    09_26, 092026...). Se não houver nenhuma, a pasta a criar é MM_AAAA."""
    pasta_ano = pasta_empresa / str(ano)
    padrao = re.compile(rf"^0?{mes}[ _-]?(?:{ano}|{ano % 100:02d})(?!\d)", re.I)
    exato = f"{mes:02d}_{ano}"
    if pasta_ano.is_dir():
        achadas = sorted(d.name for d in os.scandir(pasta_ano) if d.is_dir() and padrao.match(d.name))
        if achadas:
            return pasta_ano / (exato if exato in achadas else achadas[0])
    return pasta_ano / exato


def pasta_usuario(raiz_relatorios, subpasta, usuario):
    """IMPORTAÇÃO\\{n}-{USUÁRIO} (ex.: 7-MARIA)."""
    base = Path(raiz_relatorios, *subpasta)
    for item in os.scandir(base):
        if item.is_dir() and re.fullmatch(rf"(?:\d+\s*-\s*)?{re.escape(usuario)}", item.name.strip(), re.I):
            return Path(item.path)
    raise Pendente(f"pasta do usuário {usuario} não encontrada em {base}")


def resolver_destino(classif, cfg, indice):
    if classif.codigo is not None:
        return pasta_mes_empresa(indice.pasta(classif.codigo), classif.ano, classif.mes)
    if classif.usuario:
        base = pasta_usuario(cfg["raiz_relatorios"], classif.subpasta, classif.usuario)
    else:
        base = Path(cfg["raiz_relatorios"], *classif.subpasta)
    return base / classif.competencia


# ---------------------------------------------------------------- arquivos

def sha256(caminho):
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(1 << 20), b""):
            h.update(bloco)
    return h.hexdigest()


def pronto(caminho, estabilidade_seg):
    """Sem alteração há X segundos e sem a Domínio segurando o arquivo aberto."""
    if time.time() - caminho.stat().st_mtime < estabilidade_seg:
        return False
    try:
        with open(caminho, "r+b"):
            return True
    except OSError:
        return False


def copiar_conferido(origem, pasta_destino):
    """Copia sem nunca sobrescrever: se já existir arquivo igual, não copia;
    se existir diferente (ex.: retificação), grava como _v2, _v3..."""
    pasta_destino.mkdir(parents=True, exist_ok=True)
    hash_origem = sha256(origem)
    destino, versao = pasta_destino / origem.name, 1
    while destino.exists():
        if sha256(destino) == hash_origem:
            return destino, "JA_EXISTIA"
        versao += 1
        destino = pasta_destino / f"{origem.stem}_v{versao}{origem.suffix}"

    temporario = destino.with_name(destino.name + ".part")
    shutil.copy2(origem, temporario)
    if sha256(temporario) != hash_origem:
        temporario.unlink(missing_ok=True)
        raise OSError("hash da cópia diferente do original")
    os.replace(temporario, destino)
    return destino, "COPIADO" if versao == 1 else f"COPIADO_V{versao}"


def mover_para(origem, pasta):
    pasta.mkdir(parents=True, exist_ok=True)
    alvo = pasta / origem.name
    if alvo.exists():
        alvo = pasta / f"{origem.stem} [{datetime.now():%Y%m%d-%H%M%S}]{origem.suffix}"
    shutil.move(str(origem), str(alvo))
    return alvo


# ---------------------------------------------------------------- ouvinte

class Ouvinte:
    def __init__(self, cfg, simulacao):
        self.cfg = cfg
        self.simulacao = simulacao
        self.entrada = Path(cfg["pasta_entrada"] or BASE.parent)  # vazio = pasta acima de .Ouvinte
        self.pendentes = self.entrada / "_PENDENTES"
        self.erros = self.entrada / "_ERROS"
        self.processados = self.entrada / "_PROCESSADOS"
        self.indice = IndiceEmpresas(cfg["raiz_empresas"])
        self.falhas = {}        # arquivo -> nº de tentativas com erro de cópia
        self.ultimo_motivo = {}  # evita repetir a mesma linha de pendente/simulação no log
        self.pendentes_em = 0
        self.limpeza_em = None

    def ciclo(self):
        if not Path(self.cfg["raiz_empresas"]).is_dir():
            log.warning("Servidor inacessível (%s). Tentando no próximo ciclo.", self.cfg["raiz_empresas"])
            return
        self.indice.atualizar()

        for arq in self._arquivos(self.entrada):
            self.processar(arq)

        intervalo = self.cfg["reprocessar_pendentes_min"] * 60
        if time.time() - self.pendentes_em >= intervalo and self.pendentes.is_dir():
            self.pendentes_em = time.time()
            self.indice.atualizar(forcar=True)  # pode ter surgido a pasta que faltava
            for arq in self._arquivos(self.pendentes):
                self.processar(arq)

        self._limpar_processados()

    def _arquivos(self, pasta):
        for item in sorted(pasta.iterdir()):
            if item.is_file() and not IGNORAR.search(item.name) and pronto(item, self.cfg["estabilidade_seg"]):
                yield item

    def processar(self, arq):
        classif = None
        try:
            gerado_em = datetime.fromtimestamp(arq.stat().st_mtime)
            classif = classificar(arq.name, gerado_em, self.cfg["usuarios"])
            pasta = resolver_destino(classif, self.cfg, self.indice)
        except (NaoReconhecido, Pendente) as e:
            return self._pendente(arq, classif, str(e))

        aviso = ""
        if classif.regra == "DAS" and "SEM MOV" in pasta.name.upper():
            aviso = "DAS gerado para competência marcada como SEM MOV"

        if self.simulacao:
            novo = "" if pasta.exists() else "pasta será criada"
            chave = (arq.name, str(pasta))
            if self.ultimo_motivo.get(arq) != chave:
                self.ultimo_motivo[arq] = chave
                registrar("SIMULADO", arq, classif, str(pasta), "; ".join(x for x in (novo, aviso) if x))
            return

        try:
            destino, status = copiar_conferido(arq, pasta)
            mover_para(arq, self.processados / f"{datetime.now():%Y-%m}")
        except OSError as e:
            n = self.falhas.get(arq.name, 0) + 1
            self.falhas[arq.name] = n
            if n >= self.cfg["tentativas_antes_de_erro"]:
                self.falhas.pop(arq.name, None)
                mover_para(arq, self.erros)
                return registrar("ERRO", arq, classif, str(pasta), f"{e} (desistiu após {n} tentativas)")
            return log.warning("Falha ao copiar %s (tentativa %d): %s", arq.name, n, e)

        self.falhas.pop(arq.name, None)
        registrar(status, arq, classif, str(destino), aviso)
        if aviso:
            registrar("AVISO", arq, classif, str(destino), aviso)
        if classif.regra.startswith("INCONSISTÊNCIAS"):
            self._separar_inconsistencias(classif.competencia)

    def _separar_inconsistencias(self, competencia):
        """Atualiza o arquivo de inconsistências de cada empresa (inconsistencias.py)."""
        try:
            gravados, pendencias = inconsistencias.gerar(competencia, self.cfg, self.indice)
        except Exception:
            return log.exception("Falha ao separar inconsistências de %s por empresa", competencia)
        log.info("Inconsistências %s separadas: %d empresa(s) atualizada(s)", competencia, len(gravados))
        for p in sorted(set(pendencias)):
            log.warning("Inconsistências %s: %s", competencia, p)

    def _pendente(self, arq, classif, motivo):
        if self.ultimo_motivo.get(arq) == motivo:
            return
        self.ultimo_motivo[arq] = motivo
        if self.simulacao:
            return registrar("SIMULADO_PENDENTE", arq, classif, "", motivo)
        if arq.parent != self.pendentes:
            novo = mover_para(arq, self.pendentes)
            self.ultimo_motivo[novo] = motivo
        registrar("PENDENTE", arq, classif, "", motivo)

    def _limpar_processados(self):
        dias = self.cfg.get("retencao_processados_dias")
        hoje = datetime.now().date()
        if not dias or self.simulacao or self.limpeza_em == hoje or not self.processados.is_dir():
            return
        self.limpeza_em = hoje
        limite = time.time() - dias * 86400
        for arq in self.processados.rglob("*"):
            if arq.is_file() and arq.stat().st_mtime < limite:
                arq.unlink()

    def rodar(self, uma_vez):
        modo = "SIMULAÇÃO (nada é copiado nem movido)" if self.simulacao else "REAL"
        log.info("Ouvinte iniciado em modo %s. Entrada: %s", modo, self.entrada)
        while True:
            try:
                self.ciclo()
            except Exception:
                log.exception("Erro inesperado no ciclo")
            if uma_vez:
                break
            time.sleep(self.cfg["intervalo_seg"])


def main():
    p = argparse.ArgumentParser(description="Ouvinte do hub de automações da Domínio")
    p.add_argument("--uma-vez", action="store_true", help="roda um ciclo e sai")
    modo = p.add_mutually_exclusive_group()
    modo.add_argument("--simular", action="store_true", help="força modo simulação")
    modo.add_argument("--real", action="store_true", help="força modo real")
    args = p.parse_args()

    cfg = carregar_config()
    simulacao = cfg.get("simulacao", True)
    if args.simular:
        simulacao = True
    elif args.real:
        simulacao = False

    configurar_log()
    ouvinte = Ouvinte(cfg, simulacao)
    # uma instância por hub: tarefa agendada + duplo clique no .exe moveriam o mesmo arquivo duas vezes
    trava = "Local\\OuvinteDominio-" + hashlib.sha256(str(ouvinte.entrada).lower().encode()).hexdigest()[:16]
    _mutex = ctypes.windll.kernel32.CreateMutexW(None, False, trava)  # liberado pelo Windows quando o processo sai
    if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        return log.info("Já existe um ouvinte rodando para %s; esta cópia foi encerrada.", ouvinte.entrada)
    if args.uma_vez:
        ouvinte.pendentes_em = 0  # num ciclo avulso, reavalia os pendentes também
    ouvinte.rodar(args.uma_vez)


if __name__ == "__main__":
    main()
