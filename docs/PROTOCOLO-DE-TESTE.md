# Protocolo de teste

Objetivo: descobrir **com números** de onde vem a lentidão e se cada ajuste melhora ou piora, em vez de confiar na impressão ("parece mais rápido").

## Regras do teste

1. **Uma mudança por vez.** Se você mudar duas coisas juntas, não dá para saber qual ajudou.
2. **Teste nas máquinas que sofrem.** Uma máquina de uso leve não reproduz o problema e só serve para validar as ferramentas.
3. **Compare o mesmo horário.** Use sempre o expediente (`-FromHour 8 -ToHour 18`).
4. **No mínimo 1 dia útil por etapa, de preferência 2 a 3.** O volume de trabalho varia de um dia para o outro.
5. **Anote o que for fora do normal:** um upload grande, uma máquina reiniciada no meio do dia, alguém de férias.
6. **Nada pesado dentro do horário medido.** O `Get-OneDriveDiag.ps1` e a varredura do `Get-OneDriveChurn.ps1` percorrem centenas de milhares de arquivos e geram carga de disco própria. Na validação, isso apareceu como disco em 100% no monitor. Rode os dois **só antes das 8h ou depois das 18h**. No expediente, só o monitor e o Churn em modo escuta (`-Hours 0 -WatchMinutes 60`), que não varre nada.

## Preparar uma máquina de teste

1. Copie a pasta do projeto para **`C:\onedrive-throttle`** (fora do OneDrive).
2. **Fora do expediente** (antes das 8h ou depois das 18h), rode uma vez, como o usuário da máquina e sem admin, o diagnóstico e a varredura do que mudou:
   ```powershell
   cd C:\onedrive-throttle
   powershell -ExecutionPolicy Bypass -File .\Get-OneDriveDiag.ps1 -Label inicial
   powershell -ExecutionPolicy Bypass -File .\Get-OneDriveChurn.ps1 -Hours 24
   ```
   O diagnóstico leva uns 5 minutos e no fim abre o Explorer apontando para o relatório. A varredura lista os arquivos modificados nas últimas 24 h (rodando depois das 18h, isso cobre o dia de trabalho).
3. Instale o monitor como **tarefa agendada**, num PowerShell **aberto como administrador**. Com **coleta central** (recomendado: você recebe os relatórios de todas as máquinas numa pasta só, sem ir até elas):
   ```powershell
   cd C:\onedrive-throttle
   powershell -ExecutionPolicy Bypass -File .\Watch-OneDrive.ps1 -InstallTask -Label teste -CollectDir 'Pasta\Subpasta'
   ```
   Sem coleta central: o mesmo comando sem `-CollectDir`. Para trocar o rótulo ou a pasta de coleta, rode o `-InstallTask` de novo (ele substitui a tarefa).

   **O que a tarefa faz:**
   - Ela se chama `onedrive-throttle-watch` e dispara **no logon de qualquer usuário** e depois **a cada 15 min**, sem prazo para acabar. Se o monitor for fechado ou morto, ele volta em até 15 min. Se ele já estiver rodando, a nova instância sai na hora sem medir e registra um `[AVISO]` no log **uma vez por dia**, não a cada 15 min.
   - Ela roda **na sessão do usuário** (grupo Usuários), nunca como SYSTEM, porque o monitor precisa ver as janelas e o OneDrive de quem está logado. Não precisa de senha.
   - O monitor abre **oculto**: a tarefa chama o `wscript.exe` com o lançador `Watch-OneDrive.vbs`, que o `-InstallTask` gera na pasta do projeto e que abre o PowerShell sem janela. Nada pisca no login, e não há janela para o usuário fechar sem querer.
   - O atalho antigo da pasta Inicializar (`-InstallStartup`), se existir, é apagado de todos os perfis, para o monitor não abrir duas vezes.
   - O `Watch-OneDrive.vbs` contém o caminho da coleta. Ele é gerado na máquina e fica fora do git.

   Sem admin, ainda dá para usar o atalho antigo: `-InstallStartup` (abre uma janela minimizada e não volta se for fechado).

   **Sobre o `-CollectDir`:**
   - Um caminho **relativo** é resolvido a partir da raiz do **OneDrive corporativo do usuário logado** (`%OneDriveCommercial%`; se vazio, `%OneDrive%`). Assim o mesmo texto funciona em qualquer máquina, mesmo com usuários diferentes. Use uma pasta de uma biblioteca compartilhada que todas as máquinas de teste sincronizam (ex.: uma pasta de TI). Também aceita caminho absoluto (`D:\...` ou `\\servidor\compartilhamento\...`).
   - **A pasta precisa existir** e estar sincronizada na máquina antes. O monitor não cria a pasta de coleta nem as pastas acima dela. Ele só cria, dentro dela, a subpasta da máquina: `<NOMEDOPC>_<USUARIO>`.
   - O monitor continua gravando tudo em `reports\` local. A coleta é uma **cópia**, feita ao iniciar e depois a cada 60 min, só dos arquivos que mudaram: CSVs do monitor, resumos do `-Report`, relatórios do diagnóstico e o log.
   - **Os relatórios do Churn nunca são copiados**, porque têm nomes de arquivos de clientes. Eles ficam só em `reports\` local. O monitor reconhece o diagnóstico pela primeira linha do relatório, não pelo nome, então um Churn com `-Label` diferente também fica de fora.
   - Se a pasta não existir, o OneDrive não estiver configurado ou uma cópia falhar, o monitor **segue medindo** normalmente, registra `[AVISO]` ou `[ERRO]` no log e tenta de novo na hora seguinte.
   - Restrinja o acesso à pasta de coleta à equipe de TI: os relatórios têm nomes de máquinas, usuários e programas.
   - A cópia de hora em hora é pequena (alguns arquivos de texto) e acontece fora da medição do disco, mas aparece como um envio do OneDrive por hora.
4. Faça logoff e login (ou, como admin, `Start-ScheduledTask onedrive-throttle-watch` para abrir já nesta sessão). **Não aparece janela nenhuma**; confira se está rodando:
   ```powershell
   # O monitor esta rodando? (mostra o PID e a linha de comando)
   Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" |
       Where-Object CommandLine -like '*Watch-OneDrive.ps1*' | Select-Object ProcessId, CommandLine
   # A tarefa existe e disparou? (LastTaskResult 0 = ok)
   Get-ScheduledTask onedrive-throttle-watch | Get-ScheduledTaskInfo
   ```
   Confira também o log `reports\watch-<PC>.log`:
   - deve ter uma linha `[INICIO]` com a hora do login, e o `cmd=` no fim dela deve mostrar `-WindowStyle Hidden -File ...Watch-OneDrive.ps1` com o rótulo e a coleta;
   - com `-CollectDir`, logo depois deve vir `[INFO] coleta central: N de N arquivo(s) copiado(s) para ...`. Se vier `[AVISO] coleta central: a pasta de coleta nao existe`, confira o caminho e se a biblioteca está sincronizada na máquina. Depois confira na pasta de coleta se apareceu a subpasta `<NOMEDOPC>_<USUARIO>`.
5. No primeiro dia, **em horário de trabalho com o usuário usando a máquina**, rode o Churn **só em modo escuta**, sem varredura:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Get-OneDriveChurn.ps1 -Hours 0 -WatchMinutes 60
   ```
   Ele escuta por 1 hora cada arquivo criado, alterado ou apagado, inclusive os temporários do Office, que somem antes de qualquer varredura. É leve: só recebe avisos do Windows e não percorre pastas. Não abre nem baixa nenhum arquivo. Os relatórios (`onedrive-churn-...txt`) têm nomes de pastas e arquivos e ficam só em `reports\`.

O monitor descobre sozinho qual ajuste está aplicado (coluna `Config`), então você não precisa trocar o `-Label` a cada etapa.

## Etapas

| Etapa | O que fazer (como admin) | Depois | Pergunta que responde |
|---|---|---|---|
| **0. Baseline** | `.\Set-OneDriveThrottle.ps1 -Action Remove` | logoff/login | Quantos minutos por dia alguma janela fica "Não respondendo", e quais programas? Os travamentos coincidem com latência alta do disco? Quem causa: OneDrive, antivírus (ESET) ou indexador? |
| **1. Prioridade** | `.\Set-OneDriveThrottle.ps1 -Action Apply` | logoff/login | A prioridade baixa reduz os minutos com travamento e a latência? |
| **2. Ajuste dirigido** | Depende do resultado da etapa 0 (ex.: `-IgnorePatterns` para arquivos que mudam sem parar, indexador, exclusão no antivírus) | logoff/login | O culpado principal foi eliminado? |
| **Contraprova** *(opcional)* | `-Action Remove` por 1 dia | logoff/login | A melhora foi do ajuste ou de uma semana mais leve? |

Depois de cada logoff/login, confira no CSV do dia: `AjusteAtivo = sim`. Se aparecer `nao`, o OneDrive não reiniciou com o ajuste novo.

**Todo dia, uma olhada rápida no log** (`reports\watch-<PC>.log`):
- uma linha `[INICIO]` por login e, se a máquina ficou ligada, `[INFO]` de hora em hora;
- `[ERRO]` ocasional não invalida o dia: o monitor pulou só aquela leitura. Muitos erros iguais, ou um `[FATAL]`, precisam ser vistos antes de continuar a etapa;
- buraco nas linhas `[INFO]` sem um `[FIM]` = o monitor foi encerrado ou a máquina desligou. Com a tarefa agendada, deve vir um `[INICIO]` novo em até 15 min (se a máquina seguiu ligada e logada). Confira as horas que faltam no CSV.

## Parar e remover o monitor

Num PowerShell **como administrador**, na pasta do projeto:

```powershell
powershell -ExecutionPolicy Bypass -File .\Watch-OneDrive.ps1 -RemoveTask
```

Isso remove a tarefa e o `Watch-OneDrive.vbs`. O monitor que já está rodando segue até o logoff; para parar na hora, encerre o processo dele (o PID vem do comando de verificação acima):

```powershell
Stop-Process -Id <PID>
```

Se a tarefa ainda existir, ele volta em até 15 min, então remova a tarefa antes. Os relatórios em `reports\` (e na pasta de coleta) ficam.

> **Não há etapa de teto de RAM.** O Windows já reduz sozinho o working set do motor de sync, e o custo real é a memória privada, que o teto não reduz. Detalhes em [COMO-FUNCIONA.md](COMO-FUNCIONA.md#2-teto-de-ram--não-usar).

## Coletar os resultados

No fim de cada etapa, na máquina de teste:

```powershell
cd C:\onedrive-throttle
powershell -ExecutionPolicy Bypass -File .\Watch-OneDrive.ps1 -Report -FromHour 8 -ToHour 18
```

**Com coleta central (`-CollectDir`):** os arquivos já estão na pasta de coleta, em `<NOMEDOPC>_<USUARIO>\`, com no máximo 1 hora de atraso, então não precisa ir até a máquina. O resumo do `-Report` entra na próxima cópia de hora em hora. Você também pode rodar o `-Report` na sua máquina, apontando para a pasta da máquina coletada: `.\Watch-OneDrive.ps1 -Report -OutDir '<pasta de coleta>\<NOMEDOPC>_<USUARIO>'`. **Os relatórios do Churn não vão para a coleta**; para analisá-los, copie-os da máquina.

**Sem coleta central:** copie a pasta **`C:\onedrive-throttle\reports`** inteira (pendrive, pasta de rede ou e-mail) e leve para análise. Ela contém:

| Arquivo | Conteúdo |
|---|---|
| `onedrive-watch-<PC>-<data>.csv` | Uma linha por minuto do dia (abre no Excel) |
| `onedrive-resumo-<data-hora>.csv` | O resumo gerado pelo `-Report` |
| `onedrive-<label>-<PC>-<data>.txt` | Os diagnósticos pontuais |
| `onedrive-churn-<PC>-<data>.txt` | Os arquivos que mudam (do `Get-OneDriveChurn.ps1`). **Só local**: nunca vai para a coleta central |
| `watch-<PC>.log` | Log do monitor: início, fim, sinal de vida a cada hora e erros com mensagem completa |
| `onedrive-watch-<PC>-<data>-anterior-<hora>.csv` | Linhas gravadas por uma versão anterior do monitor no mesmo dia (o `-Report` lê junto) |

Os arquivos de máquinas diferentes não se misturam: o nome do PC faz parte do nome do arquivo.

## Como ler o resumo

O resumo separa cada dia e configuração em três **fases** (coluna `Fase`):

- **`inicio`**: primeiros 60 min após cada início do OneDrive (login, reboot, restart). É a sincronização inicial, sempre mais pesada.
- **`resto`**: o resto do dia, que reflete o trabalho do usuário. **Compare as etapas por esta linha.**
- **`sem sync`**: minutos com o motor de sync parado. Mostra quanta carga sobra sem o OneDrive.

**O resultado que importa é `MinTrava`**: minutos com alguma janela "Não respondendo", que é o que o usuário sente. Para o disco, a métrica é a **latência** (`LatMs` e `LatP95`). Em SSD, "% ocupado" marca 100% mesmo com o disco respondendo rápido, então `Ocup%` e `Em100%` ficam só no CSV do resumo, como referência.

| Coluna | Significado | O que é bom |
|---|---|---|
| `MinTrava` | **Minutos com alguma janela "Não respondendo"** (vazio = CSV de versão sem essa medição) | **Cair** de uma etapa para a outra |
| `Trava%` | % das amostras de 5 s com janela travada | Cair |
| `CpuMed` / `CpuP95` | CPU total da máquina (média / p95 dos minutos) | Referência: CPU alta junto com trava aponta CPU |
| `LatMs` / `LatP95` | **Latência do disco** (média / p95, em ms). SSD SATA saudável: < ~5 ms. Dezenas de ms = saturado | **Cair** |
| `MinLentos` | Minutos com latência média ≥ 20 ms (ajustável com `-SlowLatMs`) | Cair |
| `Iops` | Operações de disco por segundo | Referência (quanto trabalho havia) |
| `Ocup%` / `Em100%` | % ocupado médio / % do tempo em ≥ 95% (só no CSV do resumo) | Só referência em SSD |
| `PagUso%` / `PagUsoMB` | Uso do arquivo de paginação (média / máximo) | Baixo. Subindo com `LivreMB` baixo = **falta de RAM** |
| `LivreMB` | RAM livre média | Subir |
| `PgRd` | Page Reads/s: leituras de disco por falta de página. **Inclui arquivos fora do cache**, não é só paginação | Referência. Sozinho não indica falta de RAM |
| `SyncPrivMB` | Memória privada do motor de sync: o custo real de RAM dele | Referência. Cresce com a quantidade de itens sincronizados |
| `SyncMB`, `OdMB`, `TotOdMB` | Working set (RAM física no momento) do sync, do OneDrive.exe e do total do OneDrive | Só referência: o Windows reduz esse número sozinho quando falta RAM |
| `SyncOps`, `OdOps`, `AvOps`, `IdxOps` | Operações de I/O por segundo do sync, OneDrive, antivírus (ESET + Defender) e indexador | Mostra **quem** usa o disco |
| `SyncIO`, `OdIO`, `AvIO`, `IdxIO` | O mesmo em MB/s | Idem (não vê listar pastas e ler atributos) |
| `Ativo` | Se o ajuste estava realmente em vigor no processo | `sim` nas etapas 1 e 2 |

A tabela **"Travamentos: quem ficou 'Não respondendo'"** lista os programas que travaram, em quantos minutos e por quanto tempo, aproximadamente. Se for o `explorer`, o OneDrive é suspeito direto: os ícones e o menu de contexto do OneDrive rodam dentro do Explorer.

A tabela **"Travamento x carga"** compara, lado a lado, os minutos **com trava** e **sem trava**: latência, CPU, operações do OneDrive, do antivírus e do indexador, memória. Ela responde se o travamento vem de carga (disco ou CPU cheios) ou de espera (o programa esperando o OneDrive ou a rede, com a máquina folgada).

A tabela **"Nos minutos lentos: quem fazia I/O"** é a mais importante quando os travamentos acompanham a latência. Ela mostra quem estava usando o disco exatamente quando a latência subiu, incluindo processos fora da lista (`TopFora` por MB/s, `TopOps` por operações). Olhe as colunas `*Ops` além das `*IO`: varrer pastas e ler atributos quase não aparece em MB/s.

A tabela **"Memória: top processos"** mostra quem mais ocupa memória privada no dia. Se o arquivo de paginação estiver enchendo, é ela que diz quem consome a RAM. Pode ser o OneDrive ou outro programa.

## Regras de decisão

- **Na fase `resto`, `MinTrava` caiu claramente na etapa 1** (de preferência junto com `LatMs`/`LatP95`) → manter a prioridade e aplicar nas outras máquinas. Latência menor **sem** queda de `MinTrava` não resolve o problema do usuário.
- **Na tabela "Travamento x carga", latência e CPU são iguais com e sem trava** → o travamento não é disco nem CPU cheios. O programa espera outra coisa (OneDrive ou rede): prioridade de I/O não vai resolver. Olhe quais programas travam e o Churn (arquivos que o OneDrive mexe sem parar).
- **Os minutos com trava têm latência ou CPU claramente maiores** → o travamento acompanha a carga; siga a tabela dos minutos lentos para achar quem gera a carga.
- **Só a fase `inicio` é lenta** → o problema é a sincronização inicial (reinícios do OneDrive, quantidade de itens), não o trabalho do dia.
- **Nos minutos lentos, o antivírus ou o indexador aparecem com operações comparáveis às do OneDrive** → a etapa 2 ataca esses dois. Também vale se a fase `sem sync` já tiver minutos lentos.
- **O Churn mostra arquivos temporários ou de trava mudando sem parar** → `-IgnorePatterns` para esses nomes.
- **`PagUso%` subindo e `LivreMB` baixo o dia todo** → o gargalo é memória: veja a tabela de memória para saber quem consome e avalie upgrade de RAM ou reduzir os itens sincronizados. `PgRd` alto com RAM livre sobrando **não** é falta de RAM, são leituras de arquivo.
- **O diagnóstico mostra mais de 300 mil itens sincronizados** → acima do recomendado pela Microsoft. Nenhum ajuste de prioridade resolve a RAM do sync; entra na discussão estrutural (ver [COMO-FUNCIONA.md](COMO-FUNCIONA.md#limitações-conhecidas)).

## Aplicar nas outras máquinas

Só depois que o ajuste for validado em pelo menos uma máquina de uso pesado:
- **Uma a uma:** `Set-OneDriveThrottle.ps1 -Action Apply` (com os parâmetros validados) como admin.
- **Em lote (AD):** GPO de inicialização do computador rodando o mesmo comando (já roda como SYSTEM).
- **O monitor em mais máquinas:** `Watch-OneDrive.ps1 -InstallTask` uma vez por máquina, como admin (também pode ir num script de inicialização do computador por GPO, que roda como SYSTEM: a tarefa continua rodando na sessão do usuário).
