# onedrive-throttle

Deixa o OneDrive "educado" no Windows 10/11: menos disputa de CPU e disco, menos espaço ocupado e menos banda — **sem nenhum script rodando em segundo plano**.

Aplica uma vez, como administrador, e o próprio Windows reaplica a cada boot, login, atualização ou crash do OneDrive.

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

Compare principalmente: *Disco ocupado*, *Page Reads/s* e a linha do `OneDrive`. Para comparar dias, use o monitor contínuo abaixo. Se o total de itens passar de **300 mil**, está acima do recomendado pela Microsoft (ver [docs/COMO-FUNCIONA.md](docs/COMO-FUNCIONA.md)).

## Monitor contínuo (comparar dias)

Uma medição de 3 minutos depende muito do momento. `Watch-OneDrive.ps1` roda o dia todo em prioridade baixa e grava **uma linha por minuto** em `reports\onedrive-watch-<PC>-<data>.csv`, ao lado dos scripts (separador e decimal do Windows, abre direto no Excel). Não altera nada.

```powershell
.\Watch-OneDrive.ps1 -Label base            # deixe a janela minimizada; Ctrl+C para parar
.\Watch-OneDrive.ps1 -Report -FromHour 8 -ToHour 18   # resumo por dia e configuração
.\Watch-OneDrive.ps1 -InstallStartup -Label base      # abre sozinho (minimizado) a cada login
.\Watch-OneDrive.ps1 -RemoveStartup                   # para de abrir no login
```

Por minuto ele grava: % ocupado do disco da pasta do OneDrive (amostrado a cada 5 s, com a fração do minuto em ≥ 95%), Page Reads/s, RAM livre, RAM/disco/CPU do `OneDrive.Sync.Service`, do `OneDrive`, do Defender e do indexador, e o processo de fora dessa lista que mais fez I/O. Também grava a configuração do IFEO em vigor (`Config`) e se o motor de sync **realmente** está com ela (`AjusteAtivo`, prioridade de I/O e teto de working set lidos do processo) — o IFEO só vale depois que o OneDrive reinicia.

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

1. **HD mecânico** — a prioridade de I/O ajuda muito, mas um SSD resolve.
2. **Arquivos que mudam o tempo todo** dentro da pasta sincronizada (bancos de dados, `.pst`, arquivos de lock). O OneDrive recalcula e reenvia a cada alteração. Use `-IgnorePatterns` ou tire da pasta.
3. **Muitos itens sincronizados** (centenas de milhares). O consumo de RAM e CPU cresce com a quantidade de itens.
4. **Pastas conhecidas (Área de Trabalho/Documentos) com lixo acumulado** sincronizando.

## Licença

MIT
