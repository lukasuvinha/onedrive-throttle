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
3. Configure o monitor para iniciar sozinho no login (cria um atalho na pasta Inicializar do usuário):
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Watch-OneDrive.ps1 -InstallStartup -Label teste
   ```
   Para parar de monitorar: `.\Watch-OneDrive.ps1 -RemoveStartup`.
4. Faça logoff e login e confira se aparece uma janela do PowerShell minimizada na barra de tarefas. Não feche essa janela.
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

Depois de cada logoff/login, confira no console do monitor ou no CSV: `AjusteAtivo = sim`. Se aparecer `nao`, o OneDrive não reiniciou com o ajuste novo.

> **Não há etapa de teto de RAM.** O Windows já reduz sozinho o working set do motor de sync, e o custo real é a memória privada, que o teto não reduz. Detalhes em [COMO-FUNCIONA.md](COMO-FUNCIONA.md#2-teto-de-ram--não-usar).

## Coletar os resultados

No fim de cada etapa, na máquina de teste:

```powershell
cd C:\onedrive-throttle
powershell -ExecutionPolicy Bypass -File .\Watch-OneDrive.ps1 -Report -FromHour 8 -ToHour 18
```

Depois copie a pasta **`C:\onedrive-throttle\reports`** inteira (pendrive, pasta de rede ou e-mail) e leve para análise. Ela contém:

| Arquivo | Conteúdo |
|---|---|
| `onedrive-watch-<PC>-<data>.csv` | Uma linha por minuto do dia (abre no Excel) |
| `onedrive-resumo-<data-hora>.csv` | O resumo gerado pelo `-Report` |
| `onedrive-<label>-<PC>-<data>.txt` | Os diagnósticos pontuais |
| `onedrive-churn-<PC>-<data>.txt` | Os arquivos que mudam (do `Get-OneDriveChurn.ps1`) |
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
