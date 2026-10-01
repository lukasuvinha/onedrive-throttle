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
  I/O por processo (`*OpsS`) e top 5 de memoria privada. `-InstallStartup`/`-RemoveStartup`
  gerenciam o atalho na pasta Inicializar do usuario.
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
3. O disco em 100% é só do OneDrive ou também do Defender (`MsMpEng`) e do indexador
   (`SearchIndexer`) reagindo aos arquivos que o OneDrive mexe?
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
