# Protocolo de teste

Objetivo: descobrir **com números** de onde vem a lentidão e se cada ajuste melhora ou piora, em vez de confiar na impressão ("parece mais rápido").

## Regras do teste

1. **Uma mudança por vez.** Se você mudar duas coisas juntas, não dá para saber qual ajudou.
2. **Teste nas máquinas que sofrem.** Uma máquina de uso leve não reproduz o problema e só serve para validar as ferramentas.
3. **Compare o mesmo horário.** Use sempre o expediente (`-FromHour 8 -ToHour 18`).
4. **No mínimo 1 dia útil por etapa, de preferência 2 a 3.** O volume de trabalho varia de um dia para o outro.
5. **Anote o que for fora do normal:** um upload grande, uma máquina reiniciada no meio do dia, alguém de férias.

## Preparar uma máquina de teste

1. Copie a pasta do projeto para **`C:\onedrive-throttle`** (fora do OneDrive).
2. Rode o diagnóstico uma vez, como o usuário da máquina e sem admin:
   ```powershell
   cd C:\onedrive-throttle
   powershell -ExecutionPolicy Bypass -File .\Get-OneDriveDiag.ps1 -Label inicial
   ```
   Ele leva uns 5 minutos e no fim abre o Explorer apontando para o relatório.
3. Configure o monitor para iniciar sozinho no login:
   - `Win + R` → `shell:startup` → Enter (abre a pasta Inicializar do usuário);
   - botão direito → Novo → Atalho, com o destino:
     ```
     powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Minimized -File "C:\onedrive-throttle\Watch-OneDrive.ps1" -Label teste
     ```
   - Para parar de monitorar, apague esse atalho.
4. Faça logoff e login e confira se aparece uma janela do PowerShell minimizada na barra de tarefas. Não feche essa janela.

O monitor descobre sozinho qual ajuste está aplicado (coluna `Config`), então você não precisa trocar o `-Label` a cada etapa.

## Etapas

| Etapa | O que fazer (como admin) | Depois | Pergunta que responde |
|---|---|---|---|
| **0. Baseline** | `.\Set-OneDriveThrottle.ps1 -Action Remove` | logoff/login | Quanto tempo o disco fica em 100%? Quem causa: OneDrive, Defender ou indexador? |
| **1. Prioridade** | `.\Set-OneDriveThrottle.ps1 -Action Apply` | logoff/login | A prioridade baixa reduz o disco em 100% e a lentidão percebida? |
| **2. Ajuste dirigido** | Depende do resultado da etapa 0 (ex.: `-IgnorePatterns` para arquivos que mudam sem parar, indexador, Defender) | logoff/login | O culpado principal foi eliminado? |
| **3. Teto de RAM** *(opcional, por último)* | `-Action Apply -MaxWorkingSetMB 600` | logoff/login | A RAM do OneDrive cai sem aumentar o Page Reads/s? |
| **Contraprova** *(opcional)* | `-Action Remove` por 1 dia | logoff/login | A melhora foi do ajuste ou de uma semana mais leve? |

Depois de cada logoff/login, confira no console do monitor ou no CSV: `AjusteAtivo = sim`. Se aparecer `nao`, o OneDrive não reiniciou com o ajuste novo.

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

Os arquivos de máquinas diferentes não se misturam: o nome do PC faz parte do nome do arquivo.

## Como ler o resumo

| Coluna | Significado | O que é bom |
|---|---|---|
| `Em100%` | % do tempo em que o disco ficou ≥ 95% ocupado | Cair de uma etapa para a outra |
| `MinSat` | Minutos em que o disco passou metade do minuto ou mais em 100% | Cair |
| `Disco%` | Ocupação média do disco | Cair |
| `PgRd` / `PgRdP95` | Leituras do arquivo de paginação por segundo (média / pico) | **Não subir.** Se subir, falta RAM |
| `LivreMB` | RAM livre média | Subir |
| `SyncMB`, `OdMB`, `TotOdMB` | RAM do motor de sync, do OneDrive.exe e do total do OneDrive | Cair (no teto de RAM) |
| `SyncIO`, `OdIO`, `DefIO`, `IdxIO` | MB/s de disco do sync, OneDrive, Defender e indexador | Mostra **quem** usa o disco |
| `Ativo` | Se o ajuste estava realmente em vigor no processo | `sim` nas etapas 1 a 3 |

A segunda tabela, **"Nos minutos saturados: quem fazia I/O"**, é a mais importante da etapa 0: mostra quem estava usando o disco exatamente quando ele travou, incluindo processos fora da lista (`TopFora`).

## Regras de decisão

- **Disco em 100% caiu claramente na etapa 1** → manter a prioridade e aplicar nas outras máquinas.
- **Nos minutos saturados, Defender ou indexador aparecem com I/O comparável ao OneDrive** → a etapa 2 ataca esses dois.
- **O Churn mostra arquivos temporários ou de trava mudando sem parar** → `-IgnorePatterns` para esses nomes.
- **`PgRd` alto e `LivreMB` baixo o dia todo** → o gargalo é memória: avaliar upgrade de RAM antes do teto.
- **O teto de RAM fez `PgRd` ou `Em100%` subirem** → remover o teto (`-Action Apply` sem `-MaxWorkingSetMB`).

## Aplicar nas outras máquinas

Só depois que o ajuste for validado em pelo menos uma máquina de uso pesado:
- **Uma a uma:** `Set-OneDriveThrottle.ps1 -Action Apply` (com os parâmetros validados) como admin.
- **Em lote (AD):** GPO de inicialização do computador rodando o mesmo comando (já roda como SYSTEM).
