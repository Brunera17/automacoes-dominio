"""Ponto de entrada do AutomacoesDominio.exe.

  AutomacoesDominio.exe                       ouvinte (padrão; é o que a tarefa agendada roda)
  AutomacoesDominio.exe resumo [MM_AAAA]      resumo recibo x DAS
  AutomacoesDominio.exe inconsistencias [MM_AAAA]

Rodado sozinho (sem config.json ao lado, ex.: na Área de Trabalho), o .exe se instala
em C:\\automações\\.Ouvinte e abre o INSTALAR.bat. Rodar a versão nova do mesmo jeito atualiza.
"""
import shutil
import subprocess
import sys
from pathlib import Path

INSTALACAO = Path(r"C:\automações\.Ouvinte")  # hub = C:\automações (pasta acima)
ARQUIVOS_DO_PACOTE = ["INSTALAR.bat", "instalar.ps1", "README.md"]  # sempre atualizados; config.json nunca


def instalar():
    exe = Path(sys.executable)
    pacote = Path(sys._MEIPASS)  # arquivos embutidos pelo --add-data do empacotar.ps1
    INSTALACAO.mkdir(parents=True, exist_ok=True)
    destino = INSTALACAO / exe.name
    # o ouvinte instalado segura o .exe aberto: encerra só os processos daquele caminho
    subprocess.run(["powershell", "-NoProfile", "-Command",
                    f"Get-Process AutomacoesDominio -EA SilentlyContinue | ? {{ $_.Path -eq '{destino}' }} | Stop-Process -Force"],
                   creationflags=subprocess.CREATE_NO_WINDOW)
    shutil.copy2(exe, destino)
    for nome in ARQUIVOS_DO_PACOTE:
        shutil.copy2(pacote / nome, INSTALACAO / nome)
    if not (INSTALACAO / "config.json").exists():
        shutil.copy2(pacote / "config.json", INSTALACAO / "config.json")
    # o instalador é uma janela (Windows Forms); o console do PowerShell fica escondido
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden",
                    "-File", str(INSTALACAO / "instalar.ps1")], creationflags=subprocess.CREATE_NO_WINDOW)


if getattr(sys, "frozen", False) and not (Path(sys.executable).parent / "config.json").exists():
    instalar()
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
