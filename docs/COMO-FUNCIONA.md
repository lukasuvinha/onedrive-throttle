# Como o projeto funciona

## O problema

Em PCs Windows com OneDrive sincronizando muitos arquivos, os programas congelam ("Não respondendo") enquanto esperam o OneDrive, o disco ou a rede. Disco em 100%, RAM alta e CPU ocupada são sintomas que podem ou não estar por trás de cada congelamento; o monitor mede os dois para separar uma coisa da outra. Fechar o OneDrive "resolve", mas aí o usuário perde o acesso às pastas sincronizadas.

A lentidão quase nunca vem só do OneDrive. Cada arquivo que o usuário salva numa pasta sincronizada dispara uma reação em cadeia:

```
usuário salva arquivo
   ├─► OneDrive: lê o arquivo, calcula o hash, envia para a nuvem
   ├─► Antivírus (ESET; ou Defender): varre o arquivo em tempo real (e de novo quando o OneDrive o lê)
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

### O problema real: janelas que congelam

O que o usuário sente não é consumo alto. É **janela congelando**: Excel, Explorer ou o sistema contábil ficam "(Não respondendo)" porque estão esperando uma resposta (do disco, do OneDrive ou da rede). Por isso o monitor mede isso diretamente:

- a cada 5 s, quantas janelas estão "Não respondendo" e de quais processos;
- por minuto, em quantas das amostras houve janela travada (`TravaAmostras` de `Amostras`), o máximo de janelas travadas ao mesmo tempo (`TravaJanelasMax`) e quais processos travaram (`TravaProcs`, ex.: `EXCEL 3 / explorer 1` = Excel travado em 3 amostras, Explorer em 1);
- a **CPU total da máquina** (`CpuTotalPct` = média do minuto, `CpuTotalMax` = pico das amostras de 5 s).

O critério é o mesmo do Windows para escrever "(Não respondendo)" no título: a janela não processa mensagens há ~5 s (`IsHungAppWindow`, aplicado às janelas principais listadas pelo `Get-Process`). O monitor **não** usa `Process.Responding` por dois motivos. Ele acusa como travado um app da Loja suspenso, como o Configurações minimizado. E ele pode esperar até 5 s por janela, o que atrasaria a amostragem. A "janela fantasma" esbranquiçada que o Windows desenha por cima da janela travada pertence ao `dwm.exe` e é ignorada, para não contar cada travamento duas vezes.

Uma trava só é detectada depois de ~5 s, então engasgos curtos (1 a 4 s) não aparecem. O monitor pega os congelamentos que o usuário vê com o título "(Não respondendo)".

No `-Report`, **`MinTrava`** (minutos com alguma janela travada) e **`Trava%`** (% das amostras com trava) são as primeiras colunas. Duas tabelas ajudam a achar a causa:

- **"Quem ficou Não respondendo"**: processo, em quantos minutos travou e o tempo aproximado travado.
- **"Travamento x carga"**: compara os minutos com trava e sem trava, lado a lado (latência, CPU, operações do OneDrive, do antivírus e do indexador, memória). Se a carga sobe junto com a trava, o travamento acompanha o disco ou a CPU cheios. Se a carga fica igual, o programa está travando esperando outra coisa, como uma resposta do OneDrive (abrir ou salvar um arquivo que ele está sincronizando ou baixando) ou da rede.

### Log e resistência a erros

O monitor fica o dia todo minimizado numa máquina de usuário, onde ninguém olha a janela. Por isso ele grava um log em **`reports\watch-<PC>.log`**:

| Linha | Quando |
|---|---|
| `[INICIO]` | Ao abrir: usuário, versão do PowerShell, intervalo, rótulo, disco medido, data do script e a **linha de comando real**, que mostra se veio do atalho de inicialização |
| `[INFO]` | Uma vez por hora (sinal de vida: linhas gravadas e erros até ali) e quando um intervalo é descartado por suspensão ou hibernação |
| `[ERRO]` | Qualquer leitura ou gravação que falhou, com a **mensagem completa**: tipo do erro, mensagens internas, linha do script e pilha de chamadas |
| `[AVISO]` | O atalho foi aberto com um monitor já rodando; a segunda instância sai sem medir |
| `[FIM]` | Ao terminar: motivo, minutos rodados, linhas gravadas e total de erros |
| `[FATAL]` | Erro fora do laço (ex.: na preparação) que impediu o monitor de rodar |

**O laço não para por erro.** Cada leitura de 5 s (disco, CPU, janelas) é independente: a que falhar fica vazia naquela amostra e vai para o log, e as outras continuam. A detecção de travamento, por exemplo, segue mesmo se o contador de disco falhar. Uma amostra em que a leitura de janelas falhou não conta como "sem trava": fica fora de `Amostras`. Se a linha do minuto inteira falhar, só ela se perde, a base de comparação é refeita e a próxima linha sai normal. Cada consulta WMI tem limite de 30 s, então uma consulta travada vira erro em vez de congelar o monitor.

Para o log não encher, um erro que se repete (o mesmo local, tipo e linha) é gravado completo na primeira vez e depois no máximo a cada 10 min, com a contagem de repetições. Se o CSV ficar aberto no Excel, as linhas esperam na memória (até 1 dia) e são gravadas quando o arquivo for fechado.

Fechar a janela, fazer logoff ou desligar mata o processo sem chance de gravar o `[FIM]`. Nesses casos, a última linha `[INFO]` de hora em hora (ou a última linha do CSV) mostra até quando ele rodou.

### O que o monitor mede, e por que a latência é a métrica principal

Em SSD, "% ocupado" só quer dizer que havia pelo menos uma operação pendente. Na validação, o disco ficou "100% ocupado" na maior parte da manhã com latência mediana de ~5 ms, ou seja, sem gargalo real. Por isso a **métrica principal é a latência** do disco da pasta do OneDrive:

- **Latência média** (`DiscoLatMs`), ponderada pelas operações. O `-Report` mostra média, p95 e "minutos lentos" (latência ≥ 20 ms, ajustável com `-SlowLatMs`). SSD SATA saudável fica abaixo de ~5 ms; dezenas de ms são saturação de verdade.
- **IOPS** (`DiscoIops`), para saber quanto trabalho havia.
- **% ocupado** e a fração do minuto em ≥ 95% (`DiscoPct95`) continuam no CSV, só como referência.

Por processo ele grava **MB/s** e também **operações/s** (`*OpsS`). O MB/s não enxerga abrir arquivo, listar pasta e ler atributos, que são boa parte do trabalho do OneDrive, do antivírus e do indexador. Nos testes, já houve minuto com disco ocupado e nenhum processo com MB/s relevante.

O grupo **Antivírus** soma o ESET (`ekrn`, `egui`, o antivírus da empresa) e o Defender (`MsMpEng` e serviços), porque uma máquina pode ter um, o outro ou os dois. Nos CSVs antigos as colunas se chamavam `Defender*`; o `-Report` lê os dois nomes.

### Memória: Page Reads/s não é só paginação

`PageReadsS` (Page Reads/s) conta **toda leitura de disco feita por falta de página**. Isso inclui o arquivo de paginação, mas também arquivos que não estavam no cache do Windows e arquivos mapeados em memória: o cache do Windows lê arquivos por esse mesmo mecanismo. Na validação, o Page Reads/s ficou em ~1.850/s com 4 GB de RAM livre e acompanhava as IOPS do disco. Eram leituras de arquivo, não falta de RAM.

Para falta de RAM, o monitor grava o **uso do arquivo de paginação** (`PaginacaoUsoPct` e `PaginacaoUsoMB`, de `Win32_PerfRawData_PerfOS_PagingFile`), além da RAM livre e do commit. **Falta de RAM = arquivo de paginação enchendo + RAM livre baixa.** O **top 5 de memória privada** (`TopMemMB`) mostra quem consome a RAM: o OneDrive ou outro programa.

### Primeira hora depois de iniciar o OneDrive

Ao iniciar (login, reboot, atualização, crash), o motor de sync reverifica as pastas sincronizadas, e essa primeira hora costuma ser bem mais pesada que o resto do dia. O monitor grava a hora de início do processo (`SyncInicio`), e o `-Report` separa cada linha em uma fase:

- **inicio**: primeiros 60 min após cada início do OneDrive (ajustável com `-StartMinutes`);
- **resto**: o restante do dia, que reflete o trabalho do usuário;
- **sem sync**: o motor de sync não estava rodando. Serve de comparação: o que sobra de carga sem o OneDrive.

Assim, um dia com vários reinícios do OneDrive não parece pior só por causa das sincronizações iniciais. Nos CSVs antigos (sem `SyncInicio`), o início é estimado pela primeira vez que o processo aparece no CSV.

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
- **O custo real é a memória privada.** É a memória que o processo alocou e que ocupa RAM ou arquivo de paginação; o teto não reduz um byte dela. O que passa do teto vai para o arquivo de paginação e volta como **leitura de disco** quando o sync precisa dela. Ou seja, o teto troca RAM por disco, e disco é justamente o que se quer aliviar.
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

Se o monitor mostrar que o antivírus ou o indexador são parte relevante da carga:

- **Indexador:** tirar as pastas do OneDrive da indexação do Windows Search. Custo: a busca pelo *conteúdo* dos arquivos no Explorer fica lenta (a busca por nome continua funcionando).
- **Antivírus (ESET; ou Defender):** excluir os *processos* do OneDrive (não as pastas) da varredura evita varrer duas vezes o mesmo arquivo. Se o ESET for gerenciado por console (ESET PROTECT), a exclusão costuma ir na política do console, não na máquina. Custo de segurança: arquivos *baixados* da nuvem pelo OneDrive só são varridos quando alguém os abre. Decisão a tomar com dados e com cuidado.

## Limitações conhecidas

- Nenhum ajuste reduz o trabalho de **sincronizar muitos itens**. O consumo de RAM do OneDrive cresce com a quantidade de arquivos e pastas sincronizados. A solução definitiva é estrutural (menos itens sincronizados, Files On-Demand, reorganização ou migração).
- **A Microsoft recomenda no máximo 300 mil itens sincronizados**, somando todas as bibliotecas da máquina. Acima disso ela avisa que o cliente pode ter problemas de desempenho. Arquivos "somente online" contam, porque o OneDrive acompanha cada item, baixado ou não. O `Get-OneDriveDiag.ps1` mostra o total por pasta sincronizada. Bibliotecas com pastas repetidas por ano/mês passam desse limite com facilidade.
- Se a máquina compromete muito mais memória do que tem, **mais RAM** pode ser a solução mais barata e eficaz.
- `PerfOptions` no IFEO funciona desde o Windows Vista, mas não tem documentação oficial detalhada da Microsoft.
