# onedrive-throttle

Deixa o OneDrive "educado" no Windows 10/11 e **mede** se isso funcionou: menos disputa de CPU e disco, menos janelas travando à espera do OneDrive, sem precisar fechá-lo.

O **ajuste** (`Set-OneDriveThrottle.ps1`) não deixa nada rodando: é aplicado uma vez, como administrador, e o próprio Windows reaplica a cada boot, login, atualização ou crash do OneDrive. As **ferramentas de medição** só leem; a única que fica em segundo plano é o monitor, e só enquanto você quiser medir.

## O que tem neste projeto

| Arquivo | Para que serve | Altera o sistema? | Admin? |
|---|---|---|---|
| `Set-OneDriveThrottle.ps1` | Aplica/remove a prioridade baixa (CPU e disco) do OneDrive e políticas opcionais | **Sim** (registro HKLM) | Sim, para `Apply`/`Remove` |
| `Watch-OneDrive.ps1` | Monitor contínuo: 1 linha por minuto (janelas travadas, CPU, disco, OneDrive, antivírus, indexador). `-Report` resume os dias | Só a tarefa agendada que o mantém rodando (`-InstallTask`) | Só para `-InstallTask`/`-RemoveTask` |
| `Get-OneDriveDiag.ps1` | Foto da máquina: hardware, itens sincronizados, 3 min de medição | Não | Não |
| `Get-OneDriveChurn.ps1` | Quais arquivos das pastas sincronizadas estão mudando | Não | Não |
| `Uninstall-OneDriveThrottle.ps1` | Remove tudo o que o projeto instalou, em um passo | Desfaz o que os outros fizeram | Sim |
| `docs/COMO-FUNCIONA.md` | Como cada peça funciona e por quê | | |
| `docs/PROTOCOLO-DE-TESTE.md` | Passo a passo do teste em várias máquinas e como ler os resultados | | |

## Início rápido

O projeto tem **duas peças independentes**. Use uma, a outra ou as duas:

| Peça | O que faz | Liga | Desliga |
|---|---|---|---|
| **Ajuste** (`Set-OneDriveThrottle.ps1`) | Põe o OneDrive no fim da fila de CPU e disco. É a "solução" | `.\Set-OneDriveThrottle.ps1 -Action Apply` | `.\Set-OneDriveThrottle.ps1 -Action Remove` |
| **Monitor** (`Watch-OneDrive.ps1`) | Só mede (travamentos, disco, CPU), oculto, o dia todo. Serve para provar se o ajuste ajudou | `.\Watch-OneDrive.ps1 -InstallTask -Label base` | `.\Watch-OneDrive.ps1 -RemoveTask` |

Todos os comandos: PowerShell **como administrador**, dentro da pasta do projeto (se a política de execução bloquear, prefixe com `powershell -ExecutionPolicy Bypass -File`). Depois de ligar ou desligar qualquer um dos dois: **logoff/login ou reiniciar**.

**Ordem recomendada num teste:** primeiro **só o monitor**, por 1 a 2 dias (o "antes"); depois **ligue o ajuste** e deixe o monitor rodando (o "depois"); compare com `.\Watch-OneDrive.ps1 -Report -FromHour 8 -ToHour 18`. Ligar o ajuste sem o monitor também funciona; você só não terá números para comparar.

**Desligar tudo de uma vez:** `.\Uninstall-OneDriveThrottle.ps1` (desfaz os dois; detalhes em [Desinstalar](#desinstalar--remover-completamente)).

**O que já aprendemos medindo:** na maioria dos casos o problema não é o OneDrive "consumir muito", e sim os programas **esperarem** o OneDrive responder sobre arquivos e pastas. Com muitos itens sincronizados (a Microsoft recomenda no máximo 300 mil), ele demora a responder e a janela em uso congela. Por isso o resultado principal do monitor é **minutos com janela travada** (`MinTrava`), não MB de RAM.

## Por que registro em vez de script no startup

A abordagem comum é um `.bat` no startup que espera 30 s e muda a prioridade do `OneDrive.exe`. Funciona, mas:

- só pega o processo que existia naquele momento — se o OneDrive reiniciar (update, crash, troca de conta), volta ao normal;
- depende de acertar o tempo de espera;
- não mexe na prioridade de **I/O**, que é o que mais trava PC com HD mecânico.

O Windows tem um mecanismo nativo para isso: `Image File Execution Options\<exe>\PerfOptions`. O kernel lê essas chaves **quando o processo é criado** e já o inicia com a prioridade de CPU e de I/O definidas.

| | Script no startup | IFEO PerfOptions |
|---|---|---|
| Sobrevive a reboot | Sim (se rodar de novo) | Sim |
| Sobrevive a restart/update do OneDrive | Não | Sim |
| Prioridade de I/O (disco) | Não | Sim |
| Processo extra rodando | Sim | Não |
| Precisa de admin | Não | Uma vez, para gravar |

> PerfOptions não tem documentação oficial detalhada da Microsoft, mas funciona desde o Vista e é usado há anos. Por design, a prioridade de I/O via IFEO só vai até "Normal" (não dá para aumentar, só reduzir).

## O que o script faz

**Sempre** (camada de prioridade), para `OneDrive.exe`, `OneDrive.Sync.Service.exe` (motor de sync das versões 2025+), `FileCoAuth.exe`, `Microsoft.SharePoint.exe` e `OneDriveStandaloneUpdater.exe`:

- CPU: `BelowNormal` (padrão) ou `Idle`
- I/O: `Low` (padrão) ou `VeryLow`
- Memória (opcional, experimental): prioridade de página baixa

**Opcional** (políticas oficiais do OneDrive, as mesmas do ADMX/GPO):

| Parâmetro | Política | Ganho |
|---|---|---|
| `-FilesOnDemand` | `FilesOnDemandEnabled` | Disco: arquivos só baixam quando abertos |
| `-DehydrateTeamSites` | `DehydrateSyncedTeamSites` | Disco: bibliotecas SharePoint/Teams viram "somente online" |
| `-DehydrateAfterDays 30` | Storage Sense | Disco: libera arquivos não abertos há N dias |
| `-AutoUploadBandwidth` | `EnableAutomaticUploadBandwidthManagement` | Rede: só envia com banda ociosa |
| `-UploadPercent 50` | `AutomaticUploadBandwidthPercentage` | Rede: teto de % do upload |
| `-IgnorePatterns '*.pst','*.ldb'` | `EnableODIgnoreListFromGPO` | CPU/disco/rede: não sincroniza arquivos que mudam o tempo todo |

## Uso

> **Coloque a pasta do projeto fora do OneDrive** (ex.: `C:\onedrive-throttle`). Os relatórios são gravados em `reports\` dentro dela; se ela estiver no OneDrive, os próprios relatórios sincronizam e atrapalham a medição. Documentação completa em [`docs/`](docs/).

```powershell
# Ver estado atual (não precisa de admin)
.\Set-OneDriveThrottle.ps1

# Só prioridade (recomendado para começar)
.\Set-OneDriveThrottle.ps1 -Action Apply

# Pacote completo
.\Set-OneDriveThrottle.ps1 -Action Apply -FilesOnDemand -AutoUploadBandwidth `
    -IgnorePatterns '*.pst','*.ost','*.ldb','*.laccdb' -DehydrateAfterDays 30

# Desfazer tudo
.\Set-OneDriveThrottle.ps1 -Action Remove
```

Se a política de execução bloquear:

```powershell
powershell -ExecutionPolicy Bypass -File .\Set-OneDriveThrottle.ps1 -Action Apply
```

Depois de aplicar: **logoff/login ou reboot**. Confira com `-Action Status` — a coluna `PriorityClass` deve mostrar `BelowNormal`. Para ver a prioridade de I/O, use o [Process Explorer](https://learn.microsoft.com/sysinternals/downloads/process-explorer) (coluna *I/O Priority*).

## Diagnóstico antes/depois

`Get-OneDriveDiag.ps1` não altera nada: levanta hardware, pastas sincronizadas, quantidade de itens e mede por alguns minutos disco/RAM/CPU por processo, gerando um `.txt` na pasta `reports` ao lado dos scripts. Rode como o usuário (sem elevar) e **fora do horário medido** pelo monitor (antes das 8h ou depois das 18h): a contagem de itens gera carga de disco própria.

```powershell
.\Get-OneDriveDiag.ps1 -Label base          # nada aplicado
.\Get-OneDriveDiag.ps1 -Label prioridade    # depois do -Action Apply
```

Compare principalmente: *Disco ocupado*, o uso do arquivo de paginação com a RAM disponível e a linha do `OneDrive`. *Page Reads/s* inclui leituras de arquivos fora do cache, então sozinho não indica falta de RAM. Para comparar dias, use o monitor contínuo abaixo. Se o total de itens passar de **300 mil**, está acima do recomendado pela Microsoft (ver [docs/COMO-FUNCIONA.md](docs/COMO-FUNCIONA.md)).

## Monitor contínuo (comparar dias)

Uma medição de 3 minutos depende muito do momento. `Watch-OneDrive.ps1` roda o dia todo em prioridade baixa e grava **uma linha por minuto** em `reports\onedrive-watch-<PC>-<data>.csv`, ao lado dos scripts (separador e decimal do Windows, abre direto no Excel). Não altera nada.

```powershell
.\Watch-OneDrive.ps1 -Label base            # teste manual: deixe a janela minimizada; Ctrl+C para parar
.\Watch-OneDrive.ps1 -Report -FromHour 8 -ToHour 18   # resumo por dia e configuração

# Instalação (PowerShell como admin): tarefa agendada, monitor oculto na sessão de cada usuário
.\Watch-OneDrive.ps1 -InstallTask -Label base
.\Watch-OneDrive.ps1 -InstallTask -Label base -CollectDir 'Pasta\Subpasta'   # + copia de hora em hora para uma pasta central (relativa ao OneDrive do usuario)
.\Watch-OneDrive.ps1 -RemoveTask                      # remove a tarefa
```

A tarefa `onedrive-throttle-watch` dispara no logon de qualquer usuário e a cada 15 min: se o monitor for encerrado, volta sozinho, e se já estiver rodando, a nova instância sai sem medir. Roda na sessão do usuário (nunca SYSTEM), sem janela, por um lançador `.vbs` gerado na pasta do projeto. Sem admin, há o atalho antigo: `-InstallStartup` / `-RemoveStartup` (janela minimizada, não volta se for fechado). Como verificar se está rodando: [docs/PROTOCOLO-DE-TESTE.md](docs/PROTOCOLO-DE-TESTE.md).

Ele também grava um log em `reports\watch-<PC>.log` (início, fim, sinal de vida a cada hora e qualquer erro com a mensagem completa). Erro de leitura não derruba o monitor: só aquela leitura é pulada e registrada.

Por minuto ele grava, primeiro, o que o usuário sente: **janelas "Não respondendo"** (verificadas a cada 5 s: em quantas amostras houve trava e quais programas travaram) e a CPU total da máquina. Depois: latência, IOPS e % ocupado do disco da pasta do OneDrive (amostrado a cada 5 s), uso do arquivo de paginação, RAM livre, Page Reads/s, RAM/disco/CPU do `OneDrive.Sync.Service`, do `OneDrive`, do antivírus (ESET e Defender) e do indexador, e o processo de fora dessa lista que mais fez I/O. O `-Report` usa a **latência** como métrica principal e separa os primeiros 60 min após cada início do OneDrive do resto do dia. Também grava a configuração do IFEO em vigor (`Config`) e se o motor de sync **realmente** está com ela (`AjusteAtivo`, prioridade de I/O e teto de working set lidos do processo) — o IFEO só vale depois que o OneDrive reinicia.

> Usa o % ocupado **por disco**, não o `_Total`: o `_Total` é a média dos discos e esconde um disco em 100% ao lado de um ocioso. Em SSD, olhe também a latência (`DiscoLatMs`): 100% ocupado com latência baixa não é gargalo.

## Quais arquivos mudam

`Get-OneDriveChurn.ps1` mostra o que está sendo gravado nas pastas sincronizadas, por categoria (temporários, travas do Office, bancos de dados, logs...), extensão e pasta, e sugere o padrão para `-IgnorePatterns`. Lê só nome, tamanho, data e atributos; não abre nem baixa arquivos.

```powershell
.\Get-OneDriveChurn.ps1                            # varredura: modificados nas últimas 24 h (fora do expediente)
.\Get-OneDriveChurn.ps1 -Hours 0 -WatchMinutes 60  # só escuta, 1 h ao vivo (pode no expediente: não varre pastas)
```

## Distribuindo em várias máquinas (AD)

- **GPO de inicialização do computador**: Computer Configuration → Policies → Windows Settings → Scripts → Startup → PowerShell Scripts, com os parâmetros desejados. Roda como SYSTEM, então tem admin.
- **GPO nativo do OneDrive**: as políticas da tabela acima também existem no ADMX que vem com o OneDrive (`%localappdata%\Microsoft\OneDrive\<versão>\adm\`). Se preferir, use o ADMX para as políticas e o script só para a prioridade.
- **GPP Registry**: dá para replicar as chaves IFEO via Group Policy Preferences, sem script nenhum.

## O que isto NÃO faz (e por quê)

- **Não limita núcleos (afinidade).** Não existe chave de registro para isso e forçar 1 núcleo só faz a sincronização demorar mais — o OneDrive fica mais tempo ativo. Com prioridade baixa, ele já só usa CPU que estiver sobrando.
- **Não recomenda teto de RAM.** Existe `WorkingSetLimitInKB` no IFEO (parâmetro `-MaxWorkingSetMB`), mas ele só limita o working set, que o Windows já reduz sozinho quando falta RAM. O custo real é a memória privada, que o teto não reduz: o excedente vai para o arquivo de paginação e volta como leitura de disco. Detalhes em [docs/COMO-FUNCIONA.md](docs/COMO-FUNCIONA.md). O que realmente reduz a RAM do OneDrive é **sincronizar menos itens** (menos bibliotecas, excluir pastas enormes; a Microsoft recomenda no máximo 300 mil). A prioridade de página baixa só faz o Windows descartar a memória dele primeiro quando falta RAM.
- **Não usa `Idle` por padrão.** Em PC sempre ocupado, `Idle` pode deixar a sincronização parada — e arquivo não sincronizado vira conflito de versão.

## Dica: a causa raiz geralmente é outra

Se o OneDrive trava a máquina, quase sempre é um destes:

1. **Muitos itens sincronizados** (centenas de milhares). O OneDrive demora a responder ao Windows sobre cada pasta e arquivo, e a janela em uso congela esperando. Sincronizar só o necessário em cada máquina é o que mais ajuda.
2. **Arquivos que mudam o tempo todo** dentro da pasta sincronizada (bancos de dados, logs, `.pst`, arquivos de trava, programas instalados dentro dela). O OneDrive recalcula e reenvia a cada alteração. Use `-IgnorePatterns` ou tire da pasta. O `Get-OneDriveChurn.ps1` mostra quais são.
3. **Antivírus e indexador** reagindo a cada arquivo que o OneDrive mexe. O monitor separa a parte de cada um.
4. **HD mecânico ou pouca RAM.** A prioridade de I/O ajuda, mas SSD e memória resolvem.

## Desinstalar / remover completamente

Em um passo (PowerShell **como administrador**, na pasta do projeto):

```powershell
powershell -ExecutionPolicy Bypass -File .\Uninstall-OneDriveThrottle.ps1
```

Ele desfaz o ajuste de prioridade e as políticas, remove a tarefa agendada e o lançador `.vbs`, apaga o atalho antigo da pasta Inicializar de todos os perfis e encerra o monitor em execução. Os relatórios em `reports\` ficam. Opções:

| Opção | Efeito |
|---|---|
| `-DeleteReports` | Apaga também a pasta `reports\` |
| `-CollectDir 'Pasta\Subpasta'` | Apaga a cópia desta máquina (`<PC>_<USUARIO>`) na pasta de coleta central. Rode com o mesmo usuário cujo OneDrive tem a pasta |
| `-KeepThrottle` | Remove só o monitor e mantém o ajuste de prioridade |

Depois: **logoff/login ou reiniciar** (o OneDrive volta a iniciar com prioridade normal) e, se quiser, apague a pasta do projeto.

Passo a passo manual, se preferir:

1. `.\Set-OneDriveThrottle.ps1 -Action Remove` (admin). Confira com `.\Set-OneDriveThrottle.ps1`: não deve aparecer nenhuma linha de IFEO.
2. `.\Watch-OneDrive.ps1 -RemoveTask` (admin) e/ou `.\Watch-OneDrive.ps1 -RemoveStartup`.
3. Encerre o monitor: Gerenciador de Tarefas → Detalhes → `powershell.exe` cuja linha de comando tem `Watch-OneDrive.ps1` (ou reinicie).
4. Apague a pasta do projeto e, se usou coleta central, a subpasta `<PC>_<USUARIO>` na pasta de coleta.

> O `-Action Remove` apaga as políticas do OneDrive que o projeto **pode** criar (`FilesOnDemandEnabled`, limites de banda, lista de exclusão, Storage Sense). Se a sua organização define alguma delas por GPO, ela volta no próximo `gpupdate`.

## Licença

MIT
