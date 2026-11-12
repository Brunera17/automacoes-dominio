# Gera dist\AutomacoesDominio.exe (PyInstaller): um arquivo so para levar a outras maquinas.
# Na outra maquina: duplo clique no .exe -> instala em C:\automacoes\.Ouvinte e abre o INSTALAR.bat.
# Rodar nesta maquina (que tem Python + PyInstaller) sempre que mudar algum .py:
#   powershell -ExecutionPolicy Bypass -File empacotar.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$build = Join-Path $env:TEMP 'build-ouvinte'

# arquivos que o .exe solta na instalacao (config.json so se ainda nao existir la)
$dados = 'config.json', 'INSTALAR.bat', 'instalar.ps1', 'README.md' | ForEach-Object { '--add-data', "$dir\$_;." }
python -m PyInstaller --onefile --noconsole --name AutomacoesDominio --distpath "$build\dist" --workpath "$build\work" --specpath $build --log-level WARN @dados "$dir\app.py"
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller falhou.' }

New-Item -ItemType Directory -Force "$dir\dist" | Out-Null
Copy-Item "$build\dist\AutomacoesDominio.exe" "$dir\dist" -Force

# Troca tambem o .exe desta maquina. O .exe onefile roda como pai + filho e parar a
# tarefa so mata o pai, entao encerra pelo caminho (so a copia desta pasta).
$exe = Join-Path $dir 'AutomacoesDominio.exe'
Stop-ScheduledTask -TaskName 'Ouvinte Dominio' -ErrorAction SilentlyContinue
Get-Process AutomacoesDominio -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $exe } | Stop-Process -Force
Start-Sleep -Seconds 2
Copy-Item "$build\dist\AutomacoesDominio.exe" $dir -Force
Start-ScheduledTask -TaskName 'Ouvinte Dominio' -ErrorAction SilentlyContinue

Write-Host "Gerado: $dir\dist\AutomacoesDominio.exe"
