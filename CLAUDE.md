# onedrive-throttle — contexto para o Claude Code

## Objetivo
Reduzir o impacto do cliente OneDrive (CPU, RAM, disco) em PCs Windows 10 de escritório
**sem fechar o OneDrive** — os usuários precisam das pastas sincronizadas abertas o dia todo.

Metas do dono do projeto:
- Acabar com o disco em 100%.
- RAM do OneDrive idealmente ≤ ~600 MB (hoje: ~1 GB no `OneDrive.Sync.Service` + ~0,5 GB no `OneDrive.exe` logo após o login).
- Não é possível reduzir a quantidade de itens sincronizados no curto prazo (bibliotecas com muitos arquivos e pastas repetidas por ano/mês). Uma migração para SharePoint está em estudo, mas é de longo prazo.

## O que já existe
- `Set-OneDriveThrottle.ps1` — aplica/remove prioridade de CPU e I/O via IFEO `PerfOptions`,
  políticas oficiais do OneDrive (HKLM\SOFTWARE\Policies\Microsoft\OneDrive) e teto de RAM
  experimental (`WorkingSetLimitInKB`). Modos: `Status`, `Apply`, `Remove`.
- `Get-OneDriveDiag.ps1` — somente leitura. Mede hardware, pastas sincronizadas, contagem de
  itens e amostras de disco/RAM/CPU por processo. Salva `.txt` em `reports\` ao lado dos scripts.
- `Watch-OneDrive.ps1` — somente leitura. Monitor continuo: uma linha por minuto em CSV diario
  (`onedrive-watch-<PC>-<data>.csv`), com a config do IFEO e o estado real do processo de sync.
  `-Report` resume por dia/config. Usa contadores brutos (`Win32_PerfRawData_*`) e disco por
  instancia (o `_Total` e media dos discos). Grava tambem IOPS/latencia do disco, operacoes de
  I/O por processo (`*OpsS`), top 5 de memoria privada, uso do arquivo de paginacao
  (`PaginacaoUsoPct/MB`), inicio do motor de sync (`SyncInicio`), CPU total (`CpuTotalPct/Max`)
  e janelas "Nao respondendo" a cada 5 s (`Amostras`, `TravaAmostras`, `TravaJanelasMax`,
  `TravaProcs`). Travamento = `IsHungAppWindow` na janela principal do `Get-Process`; NAO usar
  `Process.Responding` (acusa app UWP suspenso e pode bloquear 5 s) e ignorar `dwm` (dono da
  "janela fantasma" de cada janela travada). Para testar trava: form WinForms +
  `[Threading.Thread]::Sleep` (o `Start-Sleep` em STA processa mensagens e nao trava).
  Log em `reports\watch-<PC>.log` (`Write-Log`: INICIO/INFO/ERRO/AVISO/FIM/FATAL; erros via
  `Write-ErrorLog`, com tipo+mensagens internas+linha+pilha e limite de 1 registro a cada 10 min
  por local/tipo/linha). Regra do laco: nenhuma leitura pode derrubar o monitor - cada leitura
  de 5 s vai em `Invoke-Safe`, a linha do minuto fica num try/catch que refaz a base, e o WMI tem
  `-OperationTimeoutSec $WmiTimeout`. Ao mexer no laco, manter isso. Teste de resistencia: copia
  no scratchpad com falhas injetadas (nao colocar ganchos de teste no script).
  `-CollectDir` (opcional): ao iniciar e a cada 60 min (`Sync-CollectDir`, junto do sinal de vida)
  copia para `<CollectDir>\<COMPUTERNAME>_<USERNAME>\` so o que mudou: onedrive-watch-*.csv,
  onedrive-resumo-*.csv, .txt do Diag e o log. Relativo = raiz do OneDrive corporativo
  (`OneDriveCommercial`, senao `OneDrive`; le tambem a variavel do usuario). Nao cria a pasta de
  coleta (so a subpasta da maquina); problema = AVISO uma vez, segue so local. NUNCA copiar
  relatorio do Churn: o Diag e identificado pela 1a linha `=== onedrive-diag |`, nao pelo nome.
  O caminho real da coleta vai so no atalho/lancador da maquina, nunca no repositorio.
  Instalacao: `-InstallTask`/`-RemoveTask` (admin) registram a tarefa `onedrive-throttle-watch`
  (grupo Usuarios pelo SID S-1-5-32-545, logon de qualquer usuario + repeticao 15 min sem fim,
  IgnoreNew, sem limite de execucao; nunca SYSTEM). Acao = wscript //B no `Watch-OneDrive.vbs`
  gerado na pasta do projeto (UTF-16, Run estilo 0, nao espera; fora do git: tem o caminho da
  coleta). Duplicata = mutex; AVISO no log 1x/dia (marcador `reports\watch-<PC>.duplicado`).
  O mutex fica logo apos o log/trap, ANTES do Add-Type (duplicata sai sem compilar C#), e so
  no modo monitor (nao em -Report/-Install*/-Remove*). `-InstallTask` aplica icacls com SIDs:
  projeto sem heranca, Admins/SYSTEM F, Usuarios RX; reports\ com Usuarios M.
  `-InstallTask` apaga o `Watch-OneDrive.lnk` da pasta Inicializar de todos os perfis.
  `-InstallStartup`/`-RemoveStartup` (atalho minimizado) ficam como alternativa sem admin.
  Atencao: `AvgDisksecPerTransfer` e `DiskTransfersPersec` brutos sao UInt32 e dao a volta
  (ver `Get-Delta32`); o formatado de latencia vem como inteiro em segundos (sempre 0).
- `Get-OneDriveChurn.ps1` — somente leitura. Quais arquivos das pastas sincronizadas mudam
  (categoria/extensao/pasta), por varredura de data (`-Hours`) e/ou escuta ao vivo com
  FileSystemWatcher (`-WatchMinutes`). So metadados; o relatorio tem nomes de arquivos e fica
  em `reports\`.

## Perfil das maquinas (importante para interpretar as medicoes)
- Todas usam SSD SATA (nenhum HD mecanico). Disco em 100% aqui significa fila de I/O
  aleatorio saturada, nao lentidao de cabeca de leitura.
- A maquina de TI (onde o projeto e desenvolvido) tem uso leve do sistema de arquivos e
  NAO apresenta disco em 100%. Serve para validar as ferramentas, nao para tirar conclusoes.
- As maquinas problematicas sao de usuarios que passam o dia criando, editando e salvando
  arquivos nas pastas sincronizadas. Cada gravacao dispara em cascata: OneDrive (hash +
  upload), Defender (varredura em tempo real) e indexador do Windows Search. A hipotese
  principal e essa soma, nao o OneDrive sozinho.
- Conclusoes e recomendacoes devem vir de medicoes nessas maquinas de uso pesado.

## Hipóteses a validar com medição (não com impressão)
1. A prioridade de I/O baixa reduz o tempo de disco em 100%? (comparar com o baseline)
2. ~~O teto de working set troca RAM por leitura de arquivo de paginação?~~ **Decidido: teto
   fora do protocolo.** O Windows ja reduz sozinho o working set do sync (validacao: ~100 MB de
   WS com ~960 MB de privada); o custo real e a memoria privada, que o teto nao reduz.
   Ver docs/COMO-FUNCIONA.md. Nao propor o teto de novo sem dado novo.
3. A latencia alta do disco e so do OneDrive ou tambem do antivirus (ESET `ekrn`/`egui`, o da
   empresa; Defender `MsMpEng` se ativo; grupo `Antivirus*` no Watch) e do indexador
   (`SearchIndexer`) reagindo aos arquivos que o OneDrive mexe?

PROBLEMA REAL (2026-10-05): janela congelando ("Nao respondendo"), nao consumo. O resultado
principal de cada etapa e `MinTrava` no -Report (fase `resto`); latencia/CPU/memoria servem para
explicar a causa (tabela "Travamento x carga": carga cheia vs. programa esperando OneDrive/rede).

Notas de medicao (validadas em 2026-10-02):
- Metrica principal de disco = latencia (`DiscoLatMs`, media e p95). Em SSD o % ocupado fica
  em 100% com latencia baixa; nao usar % ocupado para concluir nada.
- `PageReadsS` NAO e so arquivo de paginacao: inclui leitura de arquivos fora do cache e
  mapeados. Falta de RAM = `PaginacaoUsoPct/MB` subindo + `RamLivreMB` baixa.
- O -Report separa a fase `inicio` (60 min apos cada `SyncInicio`), `resto` e `sem sync`.
  Comparar etapas pela fase `resto`.
4. ~~`WorkingSetLimitInKB` via IFEO é respeitado?~~ Sem prioridade (teto fora do protocolo).
5. Quantos itens cada maquina sincroniza? A Microsoft recomenda no maximo 300 mil; acima disso
   a RAM do sync e problema estrutural, nao de ajuste.

Regra de medicao: Diag e a varredura do Churn so fora do horario medido (antes das 8h / depois
das 18h); no expediente, so o Watch e o Churn em modo escuta (`-Hours 0 -WatchMinutes N`).

## Próximo passo sugerido
Criar um monitor contínuo leve (amostra a cada 30–60 s, grava CSV diário em
`reports\` ao lado dos scripts), reaproveitando a lógica de coleta do
`Get-OneDriveDiag.ps1`, para comparar dias com e sem cada ajuste — em vez de medições de
3 minutos que dependem do momento.

## Regras de trabalho
- Relatorios (txt/csv) sempre em `reports\` ao lado dos scripts (`$PSScriptRoot`), nunca em
  pastas do sistema (AppData etc.). O usuario precisa achar e copiar facil.
- Documentacao em `docs/` (COMO-FUNCIONA.md e PROTOCOLO-DE-TESTE.md): mantenha atualizada
  quando mudar comportamento, parametros ou colunas dos scripts.
- **Pergunte antes** de alterar registro, serviços, tarefas agendadas ou configurações do
  OneDrive/Defender. Ler e medir pode sem perguntar.
- **Não abra nem leia o conteúdo de arquivos de clientes** dentro das pastas sincronizadas.
  Contar arquivos e ler metadados (tamanho, datas) é ok. Nunca force download de arquivos
  "somente online".
- O repositório é **pessoal e genérico**: nada de nome de empresa, usuário, caminho real,
  hostname ou tenant em arquivos versionados. Relatórios ficam fora do git (`.gitignore`).
- Quando pedirem para "esconder" código, **não apague**: comente o bloco no lugar com uma
  nota de onde estava e como reativar.
- Scripts devem funcionar no Windows PowerShell 5.1, sem acentos nos `.ps1` (encoding),
  e usar classes WMI `Win32_PerfFormattedData_*` em vez de `Get-Counter` (nomes de contador
  são traduzidos em Windows pt-BR).
- Explicações e mensagens em português.
