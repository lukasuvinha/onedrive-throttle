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
| `Watch-OneDrive.ps1` | Monitor contínuo: 1 linha por minuto num CSV diário. `-Report` resume os dias | Não | Não |
| `Get-OneDriveChurn.ps1` | *(em desenvolvimento)* Lista quais arquivos das pastas sincronizadas mudam, por tipo e pasta. Lê só nome, tamanho e data | Não | Não |
| `Set-OneDriveThrottle.ps1` | Aplica ou remove os ajustes (`-Action Status / Apply / Remove`) | **Sim** | Sim, para Apply e Remove |

Todos os relatórios vão para a pasta **`reports\`**, ao lado dos scripts.

> **A pasta do projeto deve ficar fora do OneDrive** (ex.: `C:\onedrive-throttle`). Se ela estiver dentro dele, os próprios relatórios sincronizam e contaminam a medição. Os scripts avisam quando isso acontece.

## Os ajustes disponíveis (`Set-OneDriveThrottle.ps1`)

### 1. Prioridade de processo (sempre aplicada no `-Action Apply`)

Grava em `HKLM\...\Image File Execution Options\<exe>\PerfOptions` a prioridade de **CPU** (BelowNormal) e de **disco** (Low) para os processos do OneDrive: `OneDrive.exe`, `OneDrive.Sync.Service.exe` (motor de sincronização, o mais pesado), `FileCoAuth.exe`, `Microsoft.SharePoint.exe` e o atualizador.

O Windows lê essa chave **quando o processo inicia**, então:
- vale a partir do próximo logoff/login (ou reboot);
- sobrevive a reboot, atualização e crash do OneDrive;
- não precisa de nenhum script rodando em segundo plano.

**Efeito:** o OneDrive continua sincronizando, mas Excel, sistema contábil e navegador passam na frente dele na fila de CPU e de disco. **Não reduz** a RAM.

### 2. Teto de RAM (opcional, experimental: `-MaxWorkingSetMB`)

Limita a RAM *física* (working set) do OneDrive. Atenção: o que passar do teto **vai para o arquivo de paginação**. Se a memória privada do processo for maior que o teto, o resultado é mais leitura de disco, ou seja, piora. Só usar se a medição mostrar que o Page Reads/s não sobe.

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
- Se a máquina compromete muito mais memória do que tem, **mais RAM** pode ser a solução mais barata e eficaz.
- `PerfOptions` no IFEO funciona desde o Windows Vista, mas não tem documentação oficial detalhada da Microsoft.
