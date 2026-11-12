"""Ponto de entrada do AutomacoesDominio.exe.

  AutomacoesDominio.exe                       ouvinte (padrão; é o que a tarefa agendada roda)
  AutomacoesDominio.exe resumo [MM_AAAA]      resumo recibo x DAS
  AutomacoesDominio.exe inconsistencias [MM_AAAA]

Rodado sozinho (sem config.json ao lado, ex.: na Área de Trabalho), o .exe se instala
em C:\\automações\\.Ouvinte e abre o INSTALAR.bat. Rodar a versão nova do mesmo jeito atualiza.
"""
import ctypes
import os
import shutil
import subprocess
import sys
from pathlib import Path

# OUVINTE_INSTALACAO só existe para testar o instalador sem mexer em C:\automações
INSTALACAO = Path(os.environ.get("OUVINTE_INSTALACAO", r"C:\automações\.Ouvinte"))  # hub = pasta acima
NOME_EXE = "AutomacoesDominio.exe"  # nome fixo instalado; as tarefas e o instalar.ps1 contam com ele
ARQUIVOS_DO_PACOTE = ["INSTALAR.bat", "instalar.ps1", "README.md"]  # sempre atualizados; config.json nunca


def aviso(texto, erro=False):
    # o .exe é --noconsole: sem isto, falhas na instalação não aparecem em lugar nenhum
    ctypes.windll.user32.MessageBoxW(None, texto, "Instalar Ouvinte Domínio", 0x10 if erro else 0x40)


def encerrar_anteriores():
    """Mata qualquer execução anterior: o ouvinte instalado (segura o .exe aberto), instaladores
    travados (.exe + PowerShell escondido) e cópias renomeadas, como 'AutomacoesDominio (1).exe'.
    Nunca este processo nem o pai (o .exe onefile roda como pai + filho)."""
    q = lambda s: str(s).replace("'", "''")  # aspas simples dentro do texto do PowerShell
    comando = f"""
$poupar = @($PID, {os.getpid()}, {os.getppid()})
Get-CimInstance Win32_Process | Where-Object {{
    $_.ProcessId -notin $poupar -and (
        $_.ExecutablePath -eq '{q(INSTALACAO / NOME_EXE)}' -or
        $_.ExecutablePath -eq '{q(sys.executable)}' -or
        $_.Name -like 'AutomacoesDominio*.exe' -or
        ($_.Name -eq 'powershell.exe' -and $_.CommandLine -like '*{q(INSTALACAO / "instalar.ps1")}*'))
}} | ForEach-Object {{ Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }}
"""
    subprocess.run(["powershell", "-NoProfile", "-Command", comando], creationflags=subprocess.CREATE_NO_WINDOW)


def instalar():
    exe = Path(sys.executable)
    pacote = Path(sys._MEIPASS)  # arquivos embutidos pelo --add-data do empacotar.ps1
    INSTALACAO.mkdir(parents=True, exist_ok=True)
    encerrar_anteriores()
    shutil.copy2(exe, INSTALACAO / NOME_EXE)  # o nome baixado pode variar; o instalado não
    for nome in ARQUIVOS_DO_PACOTE:
        shutil.copy2(pacote / nome, INSTALACAO / nome)
    if not (INSTALACAO / "config.json").exists():
        shutil.copy2(pacote / "config.json", INSTALACAO / "config.json")
    # o instalador é uma janela (Windows Forms); o console do PowerShell fica escondido
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden",
                        "-File", str(INSTALACAO / "instalar.ps1")],
                       capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
    if r.returncode not in (0, 2):  # 2 = o próprio instalador já mostrou o erro
        detalhe = r.stderr.decode("cp850", "replace").strip()[-800:] or f"código {r.returncode}"
        aviso("O instalador não conseguiu abrir. Pode ser bloqueio de antivírus ou política do PowerShell.\n\n"
              + detalhe, erro=True)


if getattr(sys, "frozen", False) and not (Path(sys.executable).parent / "config.json").exists():
    try:
        instalar()
    except Exception as e:  # ex.: .exe instalado em uso, sem permissão em C:\
        aviso(f"Falha ao instalar em {INSTALACAO}:\n\n{e}", erro=True)
    sys.exit()

comando = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("-") else "ouvinte"
if len(sys.argv) > 1 and sys.argv[1] == comando:
    del sys.argv[1]  # cada main() lê o resto de sys.argv como se fosse chamado direto

if comando == "resumo":
    import resumo
    resumo.main()
elif comando == "inconsistencias":
    import inconsistencias
    inconsistencias.main()
elif comando == "ouvinte":
    import ouvinte
    ouvinte.main()
else:
    sys.exit(f"Comando desconhecido: {comando}. Use ouvinte, resumo ou inconsistencias.")
