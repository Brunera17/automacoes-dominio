# Instala o ouvinte do hub Domínio nesta máquina (aberto pelo .exe ou pelo INSTALAR.bat).
# Janela: detecta a pasta do servidor em qualquer unidade (X:, Z:...), grava o config.json e cria as tarefas.
# Idempotente: rode de novo para trocar o destino ou depois de trocar o .exe.
#   -Detectar  só lista as pastas encontradas (diagnóstico), sem abrir janela nem instalar.
# Salvo em UTF-8 com BOM: o PowerShell 5.1 lê .ps1 sem BOM como ANSI e estraga os acentos.
param([switch]$Detectar)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
# qualquer erro não tratado vira mensagem na tela (o PowerShell roda escondido)
trap { if (-not $Detectar) { [Windows.Forms.MessageBox]::Show("Falha no instalador:`n`n$($_.Exception.Message)", 'Instalar Ouvinte Domínio', 'OK', 'Error') | Out-Null; exit 2 }; throw }  # 2 = erro já mostrado (o .exe não repete)
$dir = $PSScriptRoot                 # ...\<hub>\.Ouvinte
$hub = Split-Path $dir -Parent
$exe = Join-Path $dir 'AutomacoesDominio.exe'
$cfgPath = Join-Path $dir 'config.json'
$cfg = Get-Content $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json

function Para-UNC($caminho) {
    # X:\... -> \\servidor\compartilhamento\... (a letra pode não existir quando a tarefa inicia)
    if ($caminho -match '^([A-Za-z]):\\(.*)$') {
        $raiz = (Get-PSDrive $Matches[1] -ErrorAction SilentlyContinue).DisplayRoot
        if ($raiz -like '\\*') { return (Join-Path $raiz $Matches[2]).TrimEnd('\') }
    }
    return $caminho.TrimEnd('\')
}

# Test-Path numa unidade de rede desconectada pode esperar 20-60 s; feito na thread da janela,
# congela tudo ("Não respondendo"). Por isso cada teste roda em outro processo, com tempo limite.
function Aguardar-Jobs($jobs, $segundos) {
    $fim = (Get-Date).AddSeconds($segundos)
    while (@($jobs | Where-Object State -eq 'Running').Count -and (Get-Date) -lt $fim) {
        [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 100
    }
    $saida = foreach ($j in $jobs) { if ($j.State -eq 'Completed') { Receive-Job $j -ErrorAction SilentlyContinue } }
    $jobs | Remove-Job -Force   # o que estourou o tempo é encerrado
    return $saida
}
function Pasta-Existe($caminho, $segundos = 8) {
    if (-not $caminho) { return $false }
    $r = Aguardar-Jobs @(Start-Job { param($p) Test-Path -LiteralPath $p } -ArgumentList $caminho) $segundos
    return ($r -contains $true)
}

function Encontrar-Empresas($segundos = 15) {
    # Procura a mesma pasta do config (ex.: "DP FISCAL\SIMPLES NACIONAL") na raiz de cada unidade, todas em paralelo.
    $sufixo = $cfg.raiz_empresas -replace '^\\\\[^\\]+\\[^\\]+\\', ''
    $candidatos = @($cfg.raiz_empresas) + @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { Join-Path $_.Root $sufixo })
    $jobs = foreach ($c in $candidatos) { Start-Job { param($p) if (Test-Path -LiteralPath $p) { $p } } -ArgumentList $c }
    $achados = Aguardar-Jobs @($jobs) $segundos | ForEach-Object { Para-UNC $_ }
    # mantém a ordem: o caminho atual do config primeiro
    return @($candidatos | ForEach-Object { Para-UNC $_ } | Where-Object { $achados -contains $_ } | Select-Object -Unique)
}

function Relatorios-De($empresas) {
    # relatórios na mesma posição relativa de antes (normalmente <empresas>\.RELATÓRIOS)
    if ($cfg.raiz_relatorios.StartsWith($cfg.raiz_empresas)) { return $empresas + $cfg.raiz_relatorios.Substring($cfg.raiz_empresas.Length) }
    return $cfg.raiz_relatorios
}

function Instalar($empresas, $relatorios, $simulacao) {
    if (-not (Test-Path $exe)) { throw "Não achei $exe." }
    $cfg.raiz_empresas = $empresas
    $cfg.raiz_relatorios = $relatorios
    $cfg.simulacao = $simulacao
    # UTF-8 sem BOM: o json.load do Python não aceita BOM
    [IO.File]::WriteAllText($cfgPath, ($cfg | ConvertTo-Json), (New-Object Text.UTF8Encoding $false))

    # Tarefa 1: ouvinte, no logon, sem janela, reinicia se cair
    $acao = New-ScheduledTaskAction -Execute $exe -WorkingDirectory $dir
    $gatilho = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
    $opcoes = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
    $quem = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
    Stop-ScheduledTask -TaskName 'Ouvinte Dominio' -ErrorAction SilentlyContinue
    # o .exe onefile roda como pai + filho; parar a tarefa só mata o pai, então encerra pelo caminho
    Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -eq $exe } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Register-ScheduledTask -TaskName 'Ouvinte Dominio' -Action $acao -Trigger $gatilho -Settings $opcoes -Principal $quem -Force | Out-Null

    # Tarefa 2: resumo recibo x DAS, todo dia 13 às 08:00 (Register-ScheduledTask não faz gatilho mensal)
    schtasks /Create /TN 'Resumo Recibo x DAS' /TR "`"$exe`" resumo" /SC MONTHLY /D 13 /ST 08:00 /RL LIMITED /IT /F | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Não foi possível criar a tarefa do resumo.' }
    Set-ScheduledTask -TaskName 'Resumo Recibo x DAS' -Action (New-ScheduledTaskAction -Execute $exe -Argument 'resumo' -WorkingDirectory $dir) | Out-Null
    Set-ScheduledTask -TaskName 'Resumo Recibo x DAS' -Settings (New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 1) -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries) | Out-Null

    Start-ScheduledTask -TaskName 'Ouvinte Dominio'
}

if ($Detectar) {
    "Hub: $hub"
    "Pastas das empresas encontradas:"; Encontrar-Empresas | ForEach-Object { "  $_" }
    return
}

# ---------------------------------------------------------------- janela
[Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object Windows.Forms.Form -Property @{
    Text = 'Instalar Ouvinte Domínio'; Size = '640,400'; StartPosition = 'CenterScreen'
    FormBorderStyle = 'FixedDialog'; MaximizeBox = $false; Font = New-Object Drawing.Font('Segoe UI', 9.5)
}
function Novo($tipo, $props) { $c = New-Object "Windows.Forms.$tipo" -Property $props; $form.Controls.Add($c); $c }

Novo Label @{ Text = 'Pasta onde a Domínio deve salvar os arquivos (hub):'; Location = '20,15'; AutoSize = $true } | Out-Null
Novo TextBox @{ Text = $hub; Location = '20,38'; Size = '585,25'; ReadOnly = $true } | Out-Null

Novo Label @{ Text = 'Pasta das empresas no servidor:'; Location = '20,78'; AutoSize = $true } | Out-Null
$cbEmpresas = Novo ComboBox @{ Location = '20,101'; Size = '480,25' }
$btEmpresas = Novo Button @{ Text = 'Procurar...'; Location = '510,99'; Size = '95,28' }
$stEmpresas = Novo Label @{ Location = '20,130'; Size = '585,20' }

Novo Label @{ Text = 'Pasta dos relatórios gerais:'; Location = '20,160'; AutoSize = $true } | Out-Null
$txRelatorios = Novo TextBox @{ Location = '20,183'; Size = '480,25' }
$btRelatorios = Novo Button @{ Text = 'Procurar...'; Location = '510,181'; Size = '95,28' }
$stRelatorios = Novo Label @{ Location = '20,212'; Size = '585,20' }

$ckSimulacao = Novo CheckBox @{ Text = 'Modo simulação (só registra no log, não move nenhum arquivo)'; Location = '20,245'; AutoSize = $true; Checked = [bool]$cfg.simulacao }
$stGeral = Novo Label @{ Location = '20,280'; Size = '585,22'; ForeColor = 'DimGray' }
$btInstalar = Novo Button @{ Text = 'Instalar'; Location = '400,315'; Size = '100,32' }
$btCancelar = Novo Button @{ Text = 'Cancelar'; Location = '505,315'; Size = '100,32'; DialogResult = 'Cancel' }
$form.AcceptButton = $btInstalar; $form.CancelButton = $btCancelar; $form.ActiveControl = $cbEmpresas

function Mostrar-Status($label, $caminho) {
    $label.Text = '… verificando'; $label.ForeColor = 'DimGray'
    if (Pasta-Existe $caminho) { $label.Text = '✔ Pasta encontrada'; $label.ForeColor = 'Green' }
    else { $label.Text = '✖ Pasta não encontrada ou sem resposta (o ouvinte tenta de novo sozinho, mas confira o caminho)'; $label.ForeColor = 'Firebrick' }
}
function Escolher-Pasta($inicial) {
    $d = New-Object Windows.Forms.FolderBrowserDialog -Property @{ Description = 'Escolha a pasta'; SelectedPath = $inicial }
    if ($d.ShowDialog($form) -eq 'OK') { return (Para-UNC $d.SelectedPath) }
}

# verifica só quando a pessoa para de digitar (e não a cada tecla)
$verificando = $false
$timer = New-Object Windows.Forms.Timer -Property @{ Interval = 700 }
$timer.Add_Tick({
    $timer.Stop()
    if ($script:verificando) { $timer.Start(); return }   # DoEvents pode reentrar aqui
    $script:verificando = $true
    try {
        Mostrar-Status $stEmpresas (Para-UNC $cbEmpresas.Text)
        Mostrar-Status $stRelatorios (Para-UNC $txRelatorios.Text)
    } finally { $script:verificando = $false }
})
function Agendar-Verificacao { $stEmpresas.Text = ''; $stRelatorios.Text = ''; $timer.Stop(); $timer.Start() }

$relatoriosEditado = $false
$cbEmpresas.Add_TextChanged({
    if (-not $script:relatoriosEditado) { $txRelatorios.Text = Relatorios-De (Para-UNC $cbEmpresas.Text) }
    Agendar-Verificacao
})
$txRelatorios.Add_TextChanged({ Agendar-Verificacao })
$txRelatorios.Add_KeyPress({ $script:relatoriosEditado = $true })
$btEmpresas.Add_Click({ $p = Escolher-Pasta $cbEmpresas.Text; if ($p) { $cbEmpresas.Text = $p } })
$btRelatorios.Add_Click({ $p = Escolher-Pasta $txRelatorios.Text; if ($p) { $script:relatoriosEditado = $true; $txRelatorios.Text = $p } })

$btInstalar.Add_Click({
    $empresas = Para-UNC $cbEmpresas.Text.Trim()
    $relatorios = Para-UNC $txRelatorios.Text.Trim()
    if (-not $empresas -or -not $relatorios) { [Windows.Forms.MessageBox]::Show($form, 'Informe as duas pastas.', 'Instalar', 'OK', 'Warning') | Out-Null; return }
    $form.UseWaitCursor = $true; $btInstalar.Enabled = $false; $stGeral.Text = 'Verificando a pasta do servidor...'
    $existe = Pasta-Existe $empresas 10
    $form.UseWaitCursor = $false; $btInstalar.Enabled = $true; $stGeral.Text = ''
    if (-not $existe -and [Windows.Forms.MessageBox]::Show($form, "A pasta das empresas não abre agora:`n$empresas`n`nInstalar assim mesmo?", 'Instalar', 'YesNo', 'Question') -ne 'Yes') { return }
    $form.UseWaitCursor = $true; $btInstalar.Enabled = $false; $stGeral.Text = 'Instalando...'; $form.Refresh()
    try {
        Instalar $empresas $relatorios $ckSimulacao.Checked
        $modo = if ($ckSimulacao.Checked) { 'SIMULAÇÃO (nada é movido)' } else { 'REAL' }
        [Windows.Forms.MessageBox]::Show($form, "Ouvinte instalado e rodando em modo $modo.`n`nÚltimo passo: nas rotinas automáticas da Domínio, aponte a Pasta de Destino para:`n$hub", 'Instalado', 'OK', 'Information') | Out-Null
        $form.DialogResult = 'OK'
    } catch {
        [Windows.Forms.MessageBox]::Show($form, "Falha na instalação:`n`n$($_.Exception.Message)", 'Erro', 'OK', 'Error') | Out-Null
        $stGeral.Text = ''; $btInstalar.Enabled = $true; $form.UseWaitCursor = $false
    }
})

$form.Add_Shown({
    # a janela nasce de um processo escondido e pode abrir atrás das outras
    $form.TopMost = $true; $form.Activate(); $form.TopMost = $false
    $btInstalar.Enabled = $false
    $stGeral.Text = 'Procurando a pasta do servidor nas unidades (até 15 s)...'
    $achados = Encontrar-Empresas
    $btInstalar.Enabled = $true
    $achados | ForEach-Object { [void]$cbEmpresas.Items.Add($_) }
    $cbEmpresas.Text = if ($achados) { $achados[0] } else { $cfg.raiz_empresas }
    $stGeral.Text = if ($achados.Count -gt 1) { "$($achados.Count) pastas encontradas: escolha na lista." } elseif ($achados) { '' } else { 'Pasta do servidor não encontrada automaticamente: use "Procurar...".' }
})
[void]$form.ShowDialog()
