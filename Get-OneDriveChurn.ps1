<#
.SYNOPSIS
    Mostra QUAIS arquivos das pastas sincronizadas mudam: por tipo, por categoria
    (temporario, trava do Office, banco de dados...) e por pasta. Serve para
    decidir o -IgnorePatterns do Set-OneDriveThrottle.ps1.

.DESCRIPTION
    Nao altera nada e NAO abre arquivos: le so nome, tamanho, data e atributos
    que o Windows ja devolve na listagem da pasta. Arquivos "somente online"
    nao sao baixados.

    Duas fontes:
      1) Varredura: arquivos com data de modificacao nas ultimas -Hours horas.
         Nao ve arquivos que foram criados e apagados (temporarios do Office).
      2) Escuta ao vivo (-WatchMinutes N): registra cada evento de criar,
         alterar, apagar e renomear durante N minutos. Pega os temporarios e
         mostra quais arquivos sao gravados de novo e de novo. Rode em horario
         de trabalho normal, com o usuario usando a maquina.

    Gera um .txt em reports\ ao lado do script. O relatorio contem nomes de
    pastas e arquivos: ele fica so na maquina (fora do git).

.PARAMETER Hours
    Janela da varredura (padrao 24). 0 = nao faz a varredura.

.PARAMETER WatchMinutes
    Minutos de escuta ao vivo (padrao 0 = nao escuta). Sugestao: 30 a 60.

.PARAMETER FolderDepth
    Quantos niveis abaixo da raiz sincronizada usar para agrupar pastas (padrao 2).
    Bibliotecas com pastas por ano/mes ficam mais legiveis com 1 ou 2.

.PARAMETER Top
    Linhas por tabela (padrao 15).

.PARAMETER TimeoutSec
    Tempo maximo da varredura (padrao 300).

.EXAMPLE
    .\Get-OneDriveChurn.ps1
    O que mudou nas ultimas 24 h.

.EXAMPLE
    .\Get-OneDriveChurn.ps1 -Hours 8 -WatchMinutes 60
    Varredura das ultimas 8 h e depois 1 h de escuta ao vivo.
#>

[CmdletBinding()]
param(
    [ValidateRange(0, 720)]
    [int]$Hours = 24,

    [ValidateRange(0, 600)]
    [int]$WatchMinutes = 0,

    [ValidateRange(1, 6)]
    [int]$FolderDepth = 2,

    [int]$Top = 15,

    [int]$TimeoutSec = 300,

    [string]$Label = 'churn',

    # Padrao: subpasta reports ao lado do script.
    [string]$OutDir = ''
)

# Pasta de relatorios padrao: reports\ ao lado do script. Resolvido aqui, e nao no
# valor padrao do parametro, porque no Windows PowerShell 5.1 o $PSScriptRoot fica
# vazio dentro do param() quando o script e chamado com "powershell.exe -File"
# (como faz o atalho de inicializacao).
if (-not $OutDir) {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
    $OutDir = Join-Path $scriptDir 'reports'
}

# Relatorios ficam numa pasta "reports" ao lado do script, facil de achar e copiar.
# A pasta do projeto NAO deve ficar dentro do OneDrive: os arquivos gerados seriam
# sincronizados e contaminariam a propria medicao.
$OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
foreach ($odRoot in @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer)) {
    if ($odRoot -and $OutDir.StartsWith($odRoot.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
        Write-Warning (("A pasta de relatorios esta dentro do OneDrive ({0}). Os arquivos gerados vao " +
            "sincronizar e atrapalhar a medicao. Copie a pasta do projeto para fora do OneDrive, " +
            "ex.: C:\onedrive-throttle") -f $OutDir)
        break
    }
}

$ErrorActionPreference = 'Continue'
$out = New-Object System.Collections.Generic.List[string]
function Add([string]$s = '') { $out.Add($s); Write-Host $s }

if ($Hours -eq 0 -and $WatchMinutes -eq 0) { throw 'Use -Hours e/ou -WatchMinutes.' }

$stamp = Get-Date -Format 'yyyyMMdd-HHmm'
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$file = Join-Path $OutDir ("onedrive-{0}-{1}-{2}.txt" -f $Label, $env:COMPUTERNAME, $stamp)
Write-Host "Relatorio sera salvo em: $file" -ForegroundColor Yellow

# ---------------------------------------------------------------------------
# Classificacao
# ---------------------------------------------------------------------------

# Atributos de arquivo "somente online" (nuvem): nao baixado.
$AttrOnline = 0x400000 -bor 0x1000   # RECALL_ON_DATA_ACCESS | OFFLINE

# Categoria pelo NOME (nunca pelo conteudo). A ordem importa: a primeira que bate vence.
# A coluna Padrao e o que iria no -IgnorePatterns.
$Categories = @(
    @{ Cat = 'trava do Office (~$)';      Rx = '^~\$';                                  Pat = '~$*' }
    @{ Cat = 'trava LibreOffice';         Rx = '^\.~lock\..*#$';                        Pat = '.~lock.*' }
    @{ Cat = 'temporario (~*.tmp, .tmp)'; Rx = '^~|\.(tmp|temp|~tmp)$';                 Pat = '*.tmp' }
    @{ Cat = 'trava de banco (ldb/laccdb)'; Rx = '\.(ldb|laccdb)$';                     Pat = '*.ldb, *.laccdb' }
    @{ Cat = 'banco de dados';            Rx = '\.(mdb|accdb|db|sqlite|sqlite3|fdb|gdb|dbf|sdf)$'; Pat = '' }
    @{ Cat = 'Outlook (pst/ost)';         Rx = '\.(pst|ost)$';                          Pat = '*.pst' }
    @{ Cat = 'download incompleto';       Rx = '\.(crdownload|part|partial|download)$'; Pat = '*.crdownload, *.part' }
    @{ Cat = 'sistema (Thumbs, desktop.ini)'; Rx = '^(thumbs\.db|desktop\.ini|\.ds_store)$'; Pat = '' }
    @{ Cat = 'log';                       Rx = '\.(log|log\d+)$';                       Pat = '*.log' }
)

function Get-Category([string]$Name) {
    $n = $Name.ToLowerInvariant()
    foreach ($c in $Categories) { if ($n -match $c.Rx) { return $c.Cat } }
    'normal'
}

function Get-Ext([string]$Name) {
    $e = [IO.Path]::GetExtension($Name).ToLowerInvariant()
    if ($e) { $e } else { '(sem)' }
}

# Raiz sincronizada que contem o caminho + pasta relativa cortada em FolderDepth niveis.
function Get-FolderKey([string]$Path) {
    foreach ($r in $roots) {
        if ($Path.StartsWith($r + '\', [StringComparison]::OrdinalIgnoreCase)) {
            $rel = [IO.Path]::GetDirectoryName($Path.Substring($r.Length + 1))
            $parts = @(if ($rel) { $rel.Split('\') | Select-Object -First $FolderDepth })
            return (Join-Path $r ($parts -join '\')).TrimEnd('\')
        }
    }
    [IO.Path]::GetDirectoryName($Path)
}

function Format-Size([double]$b) {
    if ($b -ge 1GB) { '{0:N1} GB' -f ($b / 1GB) } elseif ($b -ge 1MB) { '{0:N1} MB' -f ($b / 1MB) } else { '{0:N0} KB' -f ($b / 1KB) }
}

function Add-Table($rows) {
    if ($rows) { Add (($rows | Format-Table -AutoSize | Out-String -Width 300).TrimEnd()) } else { Add '  (nada)' }
}

# Uma linha por categoria, com o padrao sugerido para -IgnorePatterns.
function Get-CategoryTable($items, [string]$CountName) {
    $items | Group-Object Cat | Sort-Object Count -Descending | ForEach-Object {
        $c = $_.Name
        [pscustomobject]@{
            Categoria  = $c
            $CountName = $_.Count
            Arquivos   = @($_.Group | Select-Object -ExpandProperty Path -Unique).Count
            Padrao     = ($Categories | Where-Object { $_.Cat -eq $c } | Select-Object -First 1).Pat
        }
    }
}

# ---------------------------------------------------------------------------
# Pastas sincronizadas (mesma deteccao do Get-OneDriveDiag.ps1)
# ---------------------------------------------------------------------------

$roots = New-Object System.Collections.Generic.List[string]
Get-ChildItem 'HKCU:\Software\Microsoft\OneDrive\Accounts' -ErrorAction SilentlyContinue | ForEach-Object {
    $acc = Get-ItemProperty $_.PSPath
    if ($acc.UserFolder) { $roots.Add($acc.UserFolder) }
    $cache = Join-Path $_.PSPath 'ScopeIdToMountPointPathCache'
    if (Test-Path $cache) {
        (Get-ItemProperty $cache).PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' } |
            ForEach-Object { $roots.Add($_.Value) }
    }
}
$roots = @($roots | Where-Object { $_ -and (Test-Path $_) } | ForEach-Object { $_.TrimEnd('\') } | Sort-Object -Unique)
# Raiz dentro de outra raiz (biblioteca atalho dentro da pasta do OneDrive) seria contada duas vezes.
$roots = @($roots | Where-Object { $r = $_; -not ($roots | Where-Object { $_ -ne $r -and $r.StartsWith($_ + '\', [StringComparison]::OrdinalIgnoreCase) }) })
# Mais especifica primeiro, para Get-FolderKey.
$roots = @($roots | Sort-Object Length -Descending)

Add "=== onedrive-churn | $Label | $(Get-Date -Format 'yyyy-MM-dd HH:mm') ==="
Add ("Pastas sincronizadas: {0}" -f $roots.Count)
foreach ($r in $roots) { Add "  $r" }
Add
if (-not $roots) { Add 'Nenhuma pasta sincronizada encontrada (rode como o usuario, sem elevar).'; $out | Set-Content -Path $file -Encoding UTF8; return }

# ---------------------------------------------------------------------------
# 1) Varredura por data de modificacao
# ---------------------------------------------------------------------------

if ($Hours -gt 0) {
    Add "--- Varredura: modificados nas ultimas $Hours h (limite de $TimeoutSec s) ---"
    $since    = (Get-Date).AddHours(-$Hours)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $recent   = New-Object System.Collections.Generic.List[object]
    $total = 0; $online = 0; $partial = $false

    foreach ($r in $roots) {
        $stack = New-Object System.Collections.Generic.Stack[IO.DirectoryInfo]
        $stack.Push([IO.DirectoryInfo]$r)
        while ($stack.Count -gt 0) {
            if ((Get-Date) -gt $deadline) { $partial = $true; break }
            $dir = $stack.Pop()
            try {
                foreach ($d in $dir.EnumerateDirectories()) { $stack.Push($d) }
                # FileInfo da listagem ja traz data/tamanho/atributos: nenhum arquivo e aberto.
                foreach ($f in $dir.EnumerateFiles()) {
                    $total++
                    $isOnline = ([int]$f.Attributes -band $AttrOnline) -ne 0
                    if ($isOnline) { $online++ }
                    if ($f.LastWriteTime -ge $since) {
                        $recent.Add([pscustomobject]@{
                            Path   = $f.FullName
                            Name   = $f.Name
                            Ext    = Get-Ext $f.Name
                            Cat    = Get-Category $f.Name
                            Folder = Get-FolderKey $f.FullName
                            Size   = [double]$f.Length
                            When   = $f.LastWriteTime
                            Online = $isOnline
                        })
                    }
                }
            } catch {}
        }
        if ($partial) { break }
    }

    Add ("Arquivos vistos: {0:N0} ({1:N0} somente online) | modificados: {2:N0} ({3}){4}" -f
        $total, $online, $recent.Count, (Format-Size ($recent | Measure-Object Size -Sum).Sum),
        $(if ($partial) { '  (PARCIAL - estourou o tempo)' } else { '' }))
    Add

    Add 'Por categoria (o que nao e "normal" e candidato a -IgnorePatterns):'
    Add-Table (Get-CategoryTable $recent 'Modificados')
    Add

    Add "Por extensao (top $Top):"
    Add-Table ($recent | Group-Object Ext | Sort-Object Count -Descending | Select-Object -First $Top | ForEach-Object {
        [pscustomobject]@{ Extensao = $_.Name; Arquivos = $_.Count; Tamanho = Format-Size ($_.Group | Measure-Object Size -Sum).Sum }
    })
    Add

    Add "Por pasta (top $Top, $FolderDepth nivel(is) abaixo da raiz):"
    Add-Table ($recent | Group-Object Folder | Sort-Object Count -Descending | Select-Object -First $Top | ForEach-Object {
        [pscustomobject]@{ Arquivos = $_.Count; Tamanho = Format-Size ($_.Group | Measure-Object Size -Sum).Sum; Pasta = $_.Name }
    })
    Add

    Add "Maiores arquivos modificados (top $Top) - cada gravacao reenvia o arquivo inteiro:"
    Add-Table ($recent | Sort-Object Size -Descending | Select-Object -First $Top | ForEach-Object {
        [pscustomobject]@{ Tamanho = Format-Size $_.Size; Modificado = $_.When.ToString('dd/MM HH:mm'); Arquivo = $_.Path }
    })
    Add
}

# ---------------------------------------------------------------------------
# 2) Escuta ao vivo
# ---------------------------------------------------------------------------

if ($WatchMinutes -gt 0) {
    if (-not ('OdtChurnWatcher' -as [type])) {
        # Os eventos chegam em outra thread: o handler em C# so enfileira, o PowerShell le a fila.
        Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using System.Collections.Concurrent;
public class OdtChurnWatcher : IDisposable {
    public ConcurrentQueue<string[]> Queue = new ConcurrentQueue<string[]>();
    public int Overflows;
    List<FileSystemWatcher> list = new List<FileSystemWatcher>();
    public void Add(string path) {
        var w = new FileSystemWatcher(path);
        w.IncludeSubdirectories = true;
        w.InternalBufferSize = 65536;
        w.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size;
        w.Created += (s, e) => Queue.Enqueue(new[] { "Criado", e.FullPath });
        // A pasta tambem "muda" quando um arquivo dentro dela muda: ruido, ignora.
        w.Changed += (s, e) => { if (!Directory.Exists(e.FullPath)) Queue.Enqueue(new[] { "Alterado", e.FullPath }); };
        w.Deleted += (s, e) => Queue.Enqueue(new[] { "Apagado", e.FullPath });
        w.Renamed += (s, e) => Queue.Enqueue(new[] { "Renomeado", e.FullPath });
        w.Error   += (s, e) => Interlocked.Increment(ref Overflows);
        w.EnableRaisingEvents = true;
        list.Add(w);
    }
    public void Dispose() { foreach (var w in list) w.Dispose(); }
}
'@
    }

    Add "--- Escuta ao vivo: $WatchMinutes min ---"
    $events = New-Object System.Collections.Generic.List[object]
    $w = New-Object OdtChurnWatcher
    try {
        foreach ($r in $roots) { $w.Add($r) }
        $end = (Get-Date).AddMinutes($WatchMinutes)
        $item = $null
        while ((Get-Date) -lt $end) {
            Start-Sleep -Seconds 1
            while ($w.Queue.TryDequeue([ref]$item)) {
                $name = [IO.Path]::GetFileName($item[1])
                $events.Add([pscustomobject]@{
                    Type   = $item[0]
                    Path   = $item[1]
                    Ext    = Get-Ext $name
                    Cat    = Get-Category $name
                    Folder = Get-FolderKey $item[1]
                })
            }
            Write-Progress -Activity 'Escutando alteracoes' -Status ("{0:N0} eventos" -f $events.Count) -SecondsRemaining ([int]($end - (Get-Date)).TotalSeconds)
        }
    } finally {
        Write-Progress -Activity 'Escutando alteracoes' -Completed
        $w.Dispose()
    }

    $mins = [math]::Max(1, $WatchMinutes)
    Add ("Eventos: {0:N0} ({1:N1}/min) em {2:N0} arquivos diferentes" -f $events.Count, ($events.Count / $mins),
        @($events | Select-Object -ExpandProperty Path -Unique).Count)
    if ($w.Overflows) { Add ("(!) O buffer de eventos estourou {0} vez(es): o numero real e MAIOR." -f $w.Overflows) }
    Add (($events | Group-Object Type | Sort-Object Count -Descending | ForEach-Object { "{0}: {1:N0}" -f $_.Name, $_.Count }) -join ' | ')
    Add
    Add 'Nota: inclui alteracoes que o proprio OneDrive baixa da nuvem (edicoes de outras pessoas).'
    Add

    Add 'Por categoria:'
    Add-Table (Get-CategoryTable $events 'Eventos')
    Add

    Add "Por extensao (top $Top):"
    Add-Table ($events | Group-Object Ext | Sort-Object Count -Descending | Select-Object -First $Top | ForEach-Object {
        [pscustomobject]@{ Extensao = $_.Name; Eventos = $_.Count; Arquivos = @($_.Group | Select-Object -ExpandProperty Path -Unique).Count }
    })
    Add

    Add "Por pasta (top $Top):"
    Add-Table ($events | Group-Object Folder | Sort-Object Count -Descending | Select-Object -First $Top | ForEach-Object {
        [pscustomobject]@{ Eventos = $_.Count; Pasta = $_.Name }
    })
    Add

    Add "Arquivos gravados mais vezes (top $Top):"
    Add-Table ($events | Group-Object Path | Sort-Object Count -Descending | Select-Object -First $Top | ForEach-Object {
        [pscustomobject]@{ Eventos = $_.Count; Categoria = $_.Group[0].Cat; Arquivo = $_.Name }
    })
    Add
}

$out | Set-Content -Path $file -Encoding UTF8
Write-Host "`nRelatorio salvo em: $file" -ForegroundColor Green
