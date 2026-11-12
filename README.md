# Ouvinte – Hub de Automações Domínio

As rotinas automáticas da Domínio salvam tudo em `C:\dev\Automações-Dominio`.
O ouvinte verifica essa pasta a cada 20 segundos e copia cada arquivo para o servidor.

## Para onde vai cada arquivo

| Arquivo | Destino (`X:\DP FISCAL\SIMPLES NACIONAL\...`) |
|---|---|
| `DAS - API Integra Contador - Empresa {cód} ... - {MMAAAA}` | `{cód}-NOME\{AAAA}\{MM_AAAA}\` |
| `Declaração Mensal_Recibo DAS - ... Empresa {cód} ... - {MMAAAA}` | `{cód}-NOME\{AAAA}\{MM_AAAA}\` |
| `Entradas_` / `Saídas_` / `Serviços_{cód}_{MMAAAA}` | `{cód}-NOME\{AAAA}\{MM_AAAA}\` |
| `NFS-e ... - {cód} - ... prestados/tomados/canceladas - Guia ...` | `.RELATÓRIOS\IMPORTAÇÃO\SERVIÇO PRESTADO\|TOMADO\|CANCELADO\{MM_AAAA}\` |
| `NF-e / NFC-e / CT-e ... - {cód} - ... entradas/saídas/canceladas` | `.RELATÓRIOS\IMPORTAÇÃO\COMERCIO ENTRADA\|SAÍDA\|CANCELADO\{NF-e\|NFC-e\|CT-e}\{MM_AAAA}\` |
| `Relatorio de Inconsistências - SN IMPORTAÇÃO ... - {USUÁRIO}` | `.RELATÓRIOS\IMPORTAÇÃO\{n}-{USUÁRIO}\{MM_AAAA}\` |
| `Relatorio de Inconsistências - ... CT-e - TODOS` | `.RELATÓRIOS\IMPORTAÇÃO\CT-e\{MM_AAAA}\` |
| `Relatorio de Inconsistências - SN APURAÇÃO / GERADOR DE GUIA ...` | `.RELATÓRIOS\APURAÇÕES\{MM_AAAA}\` |
| `Relatorio de Inconsistências - ... ZERAMENTO ...` | `.RELATÓRIOS\ZERAMENTO\{MM_AAAA}\` |
| `... - Notas Lançadas` | `.RELATÓRIOS\IMPORTAÇÃO\NOTAS LANÇADAS\{MM_AAAA}\` |

**Competência:** quando vem no nome do arquivo (`MMAAAA`), é usada essa.
Quando não vem, é o **mês anterior** ao da geração do arquivo, porque as rotinas rodam do dia 1 ao 12 sobre o mês anterior.

**Pasta do mês da empresa:** se já existir uma pasta do mês (`09_2026`, `09_2026 SEM MOV`, `09-2026`, `09_26`, `092026`), o ouvinte usa ela.
Se não existir nenhuma, cria `09_2026`, e cria também a pasta do ano se faltar.

## Regras de segurança

- **Nada é sobrescrito.** Se o destino já tiver um arquivo idêntico (mesmo hash), o ouvinte não copia de novo. Se tiver um arquivo diferente, grava como `_v2`, `_v3` e assim por diante.
- A cópia é conferida por hash. Só depois disso o original sai do hub e vai para `_PROCESSADOS\AAAA-MM\`, onde fica guardado por 90 dias.
- `_PENDENTES\` recebe os arquivos que o ouvinte não sabe para onde mandar: nome desconhecido, empresa sem pasta `{cód}-` no servidor ou código duplicado. Esses arquivos são reavaliados a cada 60 minutos, então quando você cria ou corrige a pasta da empresa eles seguem sozinhos.
- `_ERROS\` recebe os arquivos que falharam na cópia 5 vezes seguidas.
- Se o servidor estiver fora do ar, o ouvinte apenas espera e tenta de novo.
- Um DAS gerado para um mês com pasta `SEM MOV` aparece como `AVISO` no log.

## Logs

- `logs\movimentos_AAAA-MM.csv`: uma linha por arquivo, com status, empresa, competência, destino e motivo. Abre direto no Excel.
- `logs\ouvinte.log`: log técnico, com um arquivo por dia.

## Instalar em outra máquina

**Não precisa de Python na outra máquina.** Tudo roda pelo `AutomacoesDominio.exe`.

1. Copie o `dist\AutomacoesDominio.exe` para a outra máquina, em qualquer pasta (por exemplo, a Área de Trabalho), e dê dois cliques nele.
   - Ele cria `C:\automações\`, que é o hub, e coloca o ouvinte inteiro em `C:\automações\.Ouvinte\`.
   - Em seguida, abre sozinho a janela de instalação.
   - Depois de instalado, o arquivo da Área de Trabalho pode ser apagado.
2. Na janela de instalação:
   - **A pasta do servidor é procurada sozinha** em todas as unidades (`X:`, `Z:`, qualquer letra). Se houver mais de uma, elas aparecem numa lista para escolher.
   - Também dá para digitar o caminho ou usar o botão **Procurar...**.
   - Cada pasta mostra **✔ encontrada** ou **✖ não encontrada**.
   - A pasta dos relatórios é preenchida a partir da pasta das empresas.
   - Caminhos com letra, como `Z:\DP FISCAL\SIMPLES NACIONAL`, são gravados como caminho de rede (`\\servidor\...`), porque a letra pode não existir quando a tarefa inicia.
   - Clique em **Instalar**. Ele grava o `config.json`, cria as tarefas "Ouvinte Dominio" e "Resumo Recibo x DAS" e já inicia o ouvinte.
   - Para trocar o destino depois, dê dois cliques em `C:\automações\.Ouvinte\INSTALAR.bat`.
   - Para diagnóstico sem janela, rode `powershell -File instalar.ps1 -Detectar`. Ele lista as pastas encontradas.
3. Nas rotinas da Domínio dessa máquina, aponte a **Pasta de Destino** para o hub, que o instalador mostra no final.

O hub é sempre a pasta acima de `.Ouvinte`, porque `pasta_entrada` fica vazio no `config.json`.
Os logs, os resumos e o estado ficam ao lado do `.exe`.

**Comandos do .exe** (o ouvinte roda sem janela; o resultado vai para `logs\` e `resumos\`):
```
AutomacoesDominio.exe                          ouvinte (é o que a tarefa roda)
AutomacoesDominio.exe --uma-vez --simular      um ciclo de teste, sem mover nada
AutomacoesDominio.exe resumo 09_2026           resumo recibo x DAS
AutomacoesDominio.exe inconsistencias 09_2026  refaz as planilhas de inconsistências
```

**Gerar o .exe a partir do repositório** (precisa de Python 3.10+ e `pip install -r requirements.txt pyinstaller`):
1. Copie o `config.exemplo.json` para `config.json` e preencha o servidor e a lista `usuarios`. Esse arquivo fica fora do repositório, porque tem dados internos.
2. Rode o `empacotar.ps1`. O `config.json` fica embutido no `.exe` como configuração padrão de uma instalação nova.

**Atualizar depois de mudar o código** (na máquina que tem Python e PyInstaller):
```
powershell -ExecutionPolicy Bypass -File empacotar.ps1
```
Esse comando gera um novo `dist\AutomacoesDominio.exe` e troca o desta máquina.
Nas outras máquinas, dê dois cliques no `.exe` novo, de qualquer pasta. Ele substitui o instalado e mantém o `config.json` daquela máquina.
Abrir o `.exe` já instalado não cria um segundo ouvinte: a cópia extra percebe que já há um rodando e se encerra.

## Como usar

```
iniciar_ouvinte.bat --uma-vez --simular   -> mostra para onde iria cada arquivo, sem mexer em nada
iniciar_ouvinte.bat --uma-vez --real      -> processa uma vez de verdade
iniciar_ouvinte.bat                       -> roda continuamente (modo do config.json)
```

Para trocar para o modo real em definitivo, mude `"simulacao": true` para `false` no `config.json`.

### Tarefa agendada "Ouvinte Dominio"

A tarefa já está cadastrada no Agendador de Tarefas. Ela inicia quando o usuário faz logon e roda o `AutomacoesDominio.exe` sem abrir janela.
Se o ouvinte cair, a tarefa reinicia em 1 minuto. Ela também não deixa rodar duas cópias ao mesmo tempo.

A tarefa usa o logon do usuário, e não a inicialização do Windows, porque as credenciais do servidor de arquivos só ficam disponíveis com o usuário logado.

```
schtasks /Query /TN "Ouvinte Dominio"        -> ver estado
schtasks /End /TN "Ouvinte Dominio" & taskkill /F /IM AutomacoesDominio.exe   -> parar
schtasks /Run /TN "Ouvinte Dominio"          -> iniciar de novo
```

Para parar, use o `taskkill` junto com o `schtasks /End`. O `.exe` roda em dois processos, e o `/End` sozinho encerra só um deles.
Depois de alterar o `config.json`, pare e inicie de novo para a mudança valer.
Mudanças nos arquivos `.py` só valem depois de rodar o `empacotar.ps1`.

## Inconsistências separadas por empresa

Cada "Relatorio de Inconsistências" da Domínio junta várias empresas. O ouvinte separa as linhas de cada uma e grava na pasta do mês da empresa:

```
{cód}-NOME\{AAAA}\{MM_AAAA}\Inconsistências Domínio - MM_AAAA.xlsx
```

- **O que entra:** todos os relatórios de inconsistências da competência, de todas as rotinas (importação, apuração, guia, zeramento).
- **Competência:** é o mês anterior à data de **Emissão** impressa no relatório.
- **Atualização:** o arquivo é refeito sempre que chega um relatório novo, e só é regravado se o conteúdo da empresa mudou.
- **Ordem das linhas:** primeiro os erros que pedem ação, depois os avisos, e por último os rotineiros. Os rotineiros ficam em cinza, com "Rotineiro = Sim" para filtrar.
- **Mensagens rotineiras:** a lista fica em `ROTINEIRAS`, no `inconsistencias.py`, e pode ser editada.
- **Planilha aberta:** se alguém estiver com a planilha aberta no Excel, a gravação é pulada e acontece no próximo relatório.

Para gerar à mão (por exemplo, para uma competência passada):

```
AutomacoesDominio.exe inconsistencias 09_2026
```

## Resumo recibo x DAS

```
AutomacoesDominio.exe resumo           -> competência do mês anterior
AutomacoesDominio.exe resumo 08_2026   -> competência específica
```

O resumo lê as pastas das empresas direto no servidor, inclusive os arquivos salvos à mão.
Os PDFs são reconhecidos pelo **conteúdo**, não pelo nome: DAS, recibo do PGDAS-D e extrato do Simples Nacional.
O resultado vai para `resumos\recibo_x_das_MM_AAAA.csv`.

A tarefa agendada **"Resumo Recibo x DAS"** roda sozinha **todo dia 13 às 08:00**, sobre a competência anterior.
Nessa execução automática, o texto do resumo também é salvo em `resumos\recibo_x_das_MM_AAAA.txt`.
Se o computador estiver desligado às 08:00, a tarefa roda assim que ele for ligado e você fizer logon.

| Status | Significado |
|---|---|
| FALTA DAS | Débito exigível maior ou igual a R$ 10,00 e nenhum DAS na pasta |
| CONFERIR | DAS sem recibo/extrato, DAS com débito zero, DAS menor que o débito ou débito em pasta `SEM MOV` |
| OK - SEM DÉBITO | Débito zero ou abaixo de R$ 10,00 (acumula para o mês seguinte) |
| OK | Débito com DAS na pasta |

## Manutenção

- **Novo usuário nas rotinas:** acrescente o nome em `usuarios` no `config.json` e crie a pasta `.RELATÓRIOS\IMPORTAÇÃO\{n}-{NOME}`.
- **Novo tipo de arquivo:** acrescente uma regra em `regras.py`. Enquanto a regra não existir, o arquivo fica em `_PENDENTES` e não se perde.
