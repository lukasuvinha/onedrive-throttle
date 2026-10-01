# Como o projeto funciona

## O problema

Em PCs Windows com OneDrive sincronizando muitos arquivos, o OneDrive pode deixar a máquina lenta: disco em 100%, RAM alta e CPU ocupada. Fechar o OneDrive "resolve", mas aí o usuário perde o acesso às pastas sincronizadas.

A lentidão quase nunca vem só do OneDrive. Cada arquivo que o usuário salva numa pasta sincronizada dispara uma reação em cadeia:

```
usuário salva arquivo
   ├─► OneDrive: lê o arquivo, calcula o hash, envia para a nuvem
   ├─► Defender: varre o arquivo em tempo real (e de novo quando o OneDrive o lê)
   └─► Windows Search: reindexa o conteúdo
```

Num SSD SATA, milhares de gravações pequenas por dia mais esses três leitores saturam a fila de disco. Somam-se a isso a memória: se a máquina compromete mais RAM do que tem, o Windows pagina para o disco, o que gera mais I/O.

O objetivo do projeto é **medir** de onde vem a carga e **aplicar só o ajuste que os números justificarem**, sem fechar o OneDrive.

## Os scripts

| Script | O que faz | Altera o sistema? | Precisa de admin? |
|---|---|---|---|
| `Get-OneDriveDiag.ps1` | Foto da máquina: hardware, pastas sincronizadas, quantidade de itens e 3 min de medição | Não | Não |
| `Watch-OneDrive.ps1` | Monitor contínuo: 1 linha por minuto num CSV diário. `-Report` resume os dias. `-InstallStartup` / `-RemoveStartup` cria ou apaga o atalho na pasta Inicializar | Não (o atalho é só um arquivo na pasta Inicializar do usuário) | Não |
| `Get-OneDriveChurn.ps1` | Mostra quais arquivos das pastas sincronizadas mudam, por categoria (temporário, trava do Office, banco...), extensão e pasta. Varredura das últimas N horas e/ou escuta ao vivo (`-WatchMinutes`). Lê só nome, tamanho, data e atributos; nunca abre nem baixa arquivos | Não | Não |
| `Set-OneDriveThrottle.ps1` | Aplica ou remove os ajustes (`-Action Status / Apply / Remove`) | **Sim** | Sim, para Apply e Remove |

Todos os relatórios vão para a pasta **`reports\`**, ao lado dos scripts.

### O que o monitor mede, e por que mais de um número de disco

Em SSD, "% ocupado" só quer dizer que havia pelo menos uma operação pendente. Um SSD pode marcar 100% e ainda responder em 0,5 ms. Por isso o monitor grava, para o disco da pasta do OneDrive:

- **% ocupado** e a fração do minuto em ≥ 95% (`DiscoPct95`);
- **IOPS** (`DiscoIops`) e **latência média** (`DiscoLatMs`). Latência subindo para dezenas de ms é saturação de verdade.

Por processo ele grava **MB/s** e também **operações/s** (`*OpsS`). O MB/s não enxerga abrir arquivo, listar pasta e ler atributos, que são boa parte do trabalho do OneDrive, do Defender e do indexador. Nos testes, já houve minuto com disco ocupado e nenhum processo com MB/s relevante.

Para a memória, grava `PageReadsS` (leituras do arquivo de paginação), RAM livre, commit e o **top 5 de memória privada** de todos os processos (`TopMemMB`). Isso mostra se quem empurra a máquina para a paginação é o OneDrive ou outro programa.

O monitor mede o disco **por disco**, nunca pelo `_Total`: o `_Total` é a média dos discos e esconde um disco em 100% ao lado de um ocioso. O `Get-OneDriveDiag.ps1` segue a mesma regra.

> **A pasta do projeto deve ficar fora do OneDrive** (ex.: `C:\onedrive-throttle`). Se ela estiver dentro dele, os próprios relatórios sincronizam e contaminam a medição. Os scripts avisam quando isso acontece.

## Os ajustes disponíveis (`Set-OneDriveThrottle.ps1`)

### 1. Prioridade de processo (sempre aplicada no `-Action Apply`)

Grava em `HKLM\...\Image File Execution Options\<exe>\PerfOptions` a prioridade de **CPU** (BelowNormal) e de **disco** (Low) para os processos do OneDrive: `OneDrive.exe`, `OneDrive.Sync.Service.exe` (motor de sincronização, o mais pesado), `FileCoAuth.exe`, `Microsoft.SharePoint.exe` e o atualizador.

O Windows lê essa chave **quando o processo inicia**, então:
- vale a partir do próximo logoff/login (ou reboot);
- sobrevive a reboot, atualização e crash do OneDrive;
- não precisa de nenhum script rodando em segundo plano.

**Efeito:** o OneDrive continua sincronizando, mas Excel, sistema contábil e navegador passam na frente dele na fila de CPU e de disco. **Não reduz** a RAM.

### 2. Teto de RAM — não usar

O parâmetro `-MaxWorkingSetMB` continua no script, mas **saiu do protocolo de teste e não é recomendado**. Ele limita só o *working set* (a parte da memória que está na RAM física naquele momento), e a validação mostrou que isso não ataca o problema:

- **O Windows já faz isso sozinho.** Quando falta RAM, ele tira páginas do motor de sync. Na validação, o working set do `OneDrive.Sync.Service` caiu para ~100 MB, enquanto a memória privada continuou em ~960 MB. O número que o Gerenciador de Tarefas mostra engana.
- **O custo real é a memória privada.** É a memória que o processo alocou e que ocupa RAM ou arquivo de paginação; o teto não reduz um byte dela. O que passa do teto vai para o arquivo de paginação e volta como **leitura de disco** (Page Reads/s) quando o sync precisa dela. Ou seja, o teto troca RAM por disco, e disco é justamente o que se quer aliviar.
- **A memória privada do sync cresce com a quantidade de itens sincronizados.** O que a reduz é sincronizar menos itens (ver *Limitações conhecidas*), não um limite no processo.

O monitor grava as duas medidas (`SyncRamMB` = working set, `SyncPrivMB` = privada). Para avaliar a RAM do OneDrive, use a privada.

### 3. Políticas oficiais do OneDrive (opcionais)

As mesmas chaves que o GPO/ADMX do OneDrive grava:

| Parâmetro | Efeito |
|---|---|
| `-FilesOnDemand` | Arquivos só baixam quando abertos (economiza espaço) |
| `-AutoUploadBandwidth` | Só envia arquivos usando banda ociosa |
| `-UploadPercent N` | Teto de % do upload |
| `-IgnorePatterns '*.tmp','~$*'` | Não sincroniza arquivos novos com esses nomes (precisa reiniciar o OneDrive) |
| `-DehydrateAfterDays N` | Storage Sense libera arquivos não abertos há N dias |

### Desfazer tudo

```powershell
.\Set-OneDriveThrottle.ps1 -Action Remove    # como admin, depois logoff/login
```

## Ajustes fora do script (dependem dos números)

Se o monitor mostrar que o Defender ou o indexador são parte relevante da carga:

- **Indexador:** tirar as pastas do OneDrive da indexação do Windows Search. Custo: a busca pelo *conteúdo* dos arquivos no Explorer fica lenta (a busca por nome continua funcionando).
- **Defender:** excluir os *processos* do OneDrive (não as pastas) da varredura evita varrer duas vezes o mesmo arquivo. Custo de segurança: arquivos *baixados* da nuvem pelo OneDrive só são varridos quando alguém os abre. Decisão a tomar com dados e com cuidado.

## Limitações conhecidas

- Nenhum ajuste reduz o trabalho de **sincronizar muitos itens**. O consumo de RAM do OneDrive cresce com a quantidade de arquivos e pastas sincronizados. A solução definitiva é estrutural (menos itens sincronizados, Files On-Demand, reorganização ou migração).
- **A Microsoft recomenda no máximo 300 mil itens sincronizados**, somando todas as bibliotecas da máquina. Acima disso ela avisa que o cliente pode ter problemas de desempenho. Arquivos "somente online" contam, porque o OneDrive acompanha cada item, baixado ou não. O `Get-OneDriveDiag.ps1` mostra o total por pasta sincronizada. Bibliotecas com pastas repetidas por ano/mês passam desse limite com facilidade.
- Se a máquina compromete muito mais memória do que tem, **mais RAM** pode ser a solução mais barata e eficaz.
- `PerfOptions` no IFEO funciona desde o Windows Vista, mas não tem documentação oficial detalhada da Microsoft.
