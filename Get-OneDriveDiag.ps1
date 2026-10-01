<#
.SYNOPSIS
    Diagnostico do impacto do OneDrive na maquina: hardware, o que esta
    sincronizado e medicao de CPU/RAM/disco por processo durante N segundos.

.DESCRIPTION
    Gera um relatorio .txt na subpasta reports ao lado do script e abre o
    Explorer apontando para ele. Nao altera nada no sistema.
    Usa as classes WMI de desempenho (Win32_PerfFormattedData_*), que tem o
    mesmo nome em Windows de qualquer idioma - Get-Counter quebra em pt-BR.

    Rode como o USUARIO da maquina (sem "Executar como administrador"), para
    que as bibliotecas sincronizadas dele aparecam.

    Fluxo de teste sugerido:
      1. Sem nada aplicado   -> .\Get-OneDriveDiag.ps1 -Label base
      2. Com prioridade      -> .\Get-OneDriveDiag.ps1 -Label prioridade
      3. Com teto de RAM     -> .\Get-OneDriveDiag.ps1 -Label teto600
    Sempre em horario de uso normal, com o OneDrive sincronizando.

.PARAMETER Seconds
    Duracao da medicao (padrao 180).

.PARAMETER Interval
    Intervalo entre amostras em segundos (padrao 3).

.PARAMETER CountTimeoutSec
    Tempo maximo gasto contando arquivos sincronizados (padrao 120).
    Contar nao baixa arquivos "somente online".

.PARAMETER Label
    Rotulo que vai no nome do arquivo (ex.: base, prioridade, teto600).
#>

[CmdletBinding()]
param(
    [int]$Seconds = 180,
    [int]$Interval = 3,
    [int]$CountTimeoutSec = 120,
    [string]$Label = 'diag',

    # Padrao: subpasta reports ao lado do script.
    [string]$OutDir = (Join-Path $PSScriptRoot 'reports')
)

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

$stamp = Get-Date -Format 'yyyyMMdd-HHmm'
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$file  = Join-Path $OutDir ("onedrive-{0}-{1}-{2}.txt" -f $Label, $env:COMPUTERNAME, $stamp)
Write-Host "Relatorio sera salvo em: $file" -ForegroundColor Yellow

# ---------------------------------------------------------------------------
Add "=== onedrive-diag | $Label | $(Get-Date -Format 'yyyy-MM-dd HH:mm') ==="
Add

# --- Hardware ---------------------------------------------------------------
Add '--- Hardware / SO ---'
$os  = Get-CimInstance Win32_OperatingSystem
$cs  = Get-CimInstance Win32_ComputerSystem
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
Add ("SO: {0} build {1}" -f $os.Caption, $os.BuildNumber)
Add ("CPU: {0} | {1} nucleos / {2} threads" -f $cpu.Name.Trim(), $cpu.NumberOfCores, $cpu.NumberOfLogicalProcessors)
Add ("RAM total: {0:N1} GB | livre agora: {1:N1} GB" -f ($cs.TotalPhysicalMemory / 1GB), ($os.FreePhysicalMemory / 1MB))
Add ("Arquivo de paginacao: {0}" -f (((Get-CimInstance Win32_PageFileUsage) | ForEach-Object { "$($_.Name) $($_.AllocatedBaseSize) MB" }) -join '; '))
try {
    Get-PhysicalDisk -ErrorAction Stop | ForEach-Object {
        Add ("Disco: {0} | {1} | {2:N0} GB" -f $_.FriendlyName, $_.MediaType, ($_.Size / 1GB))
    }
} catch { Add 'Disco: (Get-PhysicalDisk indisponivel)' }
$sd = Get-PSDrive -Name ($env:SystemDrive.TrimEnd(':'))
Add ("Espaco livre em {0}: {1:N1} GB" -f $env:SystemDrive, ($sd.Free / 1GB))
Add

# --- OneDrive ---------------------------------------------------------------
Add '--- OneDrive ---'
$od = Get-Process OneDrive -ErrorAction SilentlyContinue | Select-Object -First 1
if ($od -and $od.Path) {
    $ver = (Get-Item $od.Path).VersionInfo.FileVersion
    # Arquitetura pelo cabecalho PE do executavel
    $arch = '?'
    try {
        $fs = [IO.File]::OpenRead($od.Path); $br = New-Object IO.BinaryReader $fs
        $fs.Position = 0x3C; $pe = $br.ReadInt32(); $fs.Position = $pe + 4
        $arch = switch ($br.ReadUInt16()) { 0x8664 { '64 bits' } 0x14C { '32 bits' } 0xAA64 { 'ARM64' } default { '?' } }
        $br.Close()
    } catch {}
    Add ("Versao: {0} ({1})" -f $ver, $arch)
    Add ("Caminho: {0}" -f $od.Path)
    Add ("Prioridade atual: {0}" -f $od.PriorityClass)
} else { Add 'OneDrive.exe nao esta rodando.' }

$ifeo = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\OneDrive.exe\PerfOptions'
if (Test-Path $ifeo) {
    $v = Get-ItemProperty $ifeo
    Add ("IFEO aplicado: CPU={0} IO={1} Page={2} WSLimitKB={3}" -f $v.CpuPriorityClass, $v.IoPriority, $v.PagePriority, $v.WorkingSetLimitInKB)
} else { Add 'IFEO: nao aplicado' }

$pol = 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'
if (Test-Path $pol) {
    $p = (Get-ItemProperty $pol).PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' } |
         ForEach-Object { "$($_.Name)=$($_.Value)" }
    Add ("Politicas: {0}" -f ($p -join ', '))
}

# Pastas sincronizadas (conta pessoal/Business + bibliotecas SharePoint)
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
$roots = $roots | Where-Object { $_ -and (Test-Path $_) } | Sort-Object -Unique
Add ("Pastas sincronizadas: {0}" -f @($roots).Count)
Add

# --- Contagem de itens -------------------------------------------------------
Add "--- Itens sincronizados (limite de $CountTimeoutSec s) ---"
$deadline = (Get-Date).AddSeconds($CountTimeoutSec)
$totalF = 0; $totalD = 0; $partial = $false
foreach ($r in $roots) {
    $f = 0; $d = 0; $stack = New-Object System.Collections.Generic.Stack[string]; $stack.Push($r)
    while ($stack.Count -gt 0) {
        if ((Get-Date) -gt $deadline) { $partial = $true; break }
        $dir = $stack.Pop()
        try {
            foreach ($x in [IO.Directory]::EnumerateDirectories($dir)) { $stack.Push($x); $d++ }
            foreach ($x in [IO.Directory]::EnumerateFiles($dir))       { $f++ }
        } catch {}
    }
    $totalF += $f; $totalD += $d
    Add ("{0,10:N0} arquivos {1,8:N0} pastas  {2}" -f $f, $d, $r)
    if ($partial) { break }
}
Add ("TOTAL: {0:N0} arquivos, {1:N0} pastas{2}" -f $totalF, $totalD, $(if ($partial) { '  (PARCIAL - estourou o tempo; o real e maior)' } else { '' }))
Add

# --- Medicao -----------------------------------------------------------------
Add "--- Medicao: $Seconds s, amostra a cada $Interval s ---"
$ncpu    = $cpu.NumberOfLogicalProcessors * @(Get-CimInstance Win32_Processor).Count
$samples = New-Object System.Collections.Generic.List[object]
$sys     = New-Object System.Collections.Generic.List[object]
$end     = (Get-Date).AddSeconds($Seconds)
$n = 0
while ((Get-Date) -lt $end) {
    $n++
    Get-CimInstance Win32_PerfFormattedData_PerfProc_Process |
        Where-Object { $_.Name -ne '_Total' -and $_.Name -ne 'Idle' } |
        ForEach-Object {
            $samples.Add([pscustomobject]@{
                N    = $n
                Name = ($_.Name -replace '#\d+$', '')
                IO   = [double]$_.IODataBytesPersec
                WS   = [double]$_.WorkingSet
                Priv = [double]$_.PrivateBytes
                CPU  = [double]$_.PercentProcessorTime / $ncpu
            })
        }
    $disk = Get-CimInstance Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "Name='_Total'"
    $mem  = Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
    $sys.Add([pscustomobject]@{
        Busy      = 100 - [double]$disk.PercentIdleTime
        Queue     = [double]$disk.CurrentDiskQueueLength
        PageReads = [double]$mem.PageReadsPersec
        AvailMB   = [double]$mem.AvailableMBytes
    })
    Write-Progress -Activity 'Medindo' -SecondsRemaining ([int]($end - (Get-Date)).TotalSeconds)
    Start-Sleep -Seconds $Interval
}
Write-Progress -Activity 'Medindo' -Completed

# Soma por nome em cada amostra (junta OneDrive#1, #2...), depois media/max
$perProc = $samples | Group-Object N, Name | ForEach-Object {
    [pscustomobject]@{
        Name = $_.Group[0].Name
        IO   = ($_.Group | Measure-Object IO   -Sum).Sum
        WS   = ($_.Group | Measure-Object WS   -Sum).Sum
        Priv = ($_.Group | Measure-Object Priv -Sum).Sum
        CPU  = ($_.Group | Measure-Object CPU  -Sum).Sum
    }
} | Group-Object Name | ForEach-Object {
    $g = $_.Group
    [pscustomobject]@{
        Processo      = $_.Name
        'Disco MB/s'  = [math]::Round((($g | Measure-Object IO -Average).Average) / 1MB, 2)
        'Disco max'   = [math]::Round((($g | Measure-Object IO -Maximum).Maximum) / 1MB, 2)
        'RAM MB'      = [math]::Round((($g | Measure-Object WS -Average).Average) / 1MB)
        'RAM max'     = [math]::Round((($g | Measure-Object WS -Maximum).Maximum) / 1MB)
        'Privada MB'  = [math]::Round((($g | Measure-Object Priv -Average).Average) / 1MB)
        'CPU %'       = [math]::Round(($g | Measure-Object CPU -Average).Average, 1)
    }
}

$watch = 'OneDrive', 'OneDrive.Sync.Service', 'FileCoAuth', 'Microsoft.SharePoint', 'MsMpEng', 'SearchIndexer', 'SearchProtocolHost', 'System'
Add 'Processos de interesse (OneDrive, antivirus, indexador):'
Add (($perProc | Where-Object { $watch -contains $_.Processo } | Sort-Object 'Disco MB/s' -Descending |
      Format-Table -AutoSize | Out-String).TrimEnd())
Add
Add 'Top 10 em disco:'
Add (($perProc | Sort-Object 'Disco MB/s' -Descending | Select-Object -First 10 |
      Format-Table -AutoSize | Out-String).TrimEnd())
Add
Add 'Top 10 em RAM:'
Add (($perProc | Sort-Object 'RAM MB' -Descending | Select-Object -First 10 |
      Format-Table -AutoSize | Out-String).TrimEnd())
Add

$busy = $sys | Measure-Object Busy -Average -Maximum
$pr   = $sys | Measure-Object PageReads -Average -Maximum
$q    = $sys | Measure-Object Queue -Average -Maximum
$av   = $sys | Measure-Object AvailMB -Average -Minimum
$pct100 = [math]::Round(100 * @($sys | Where-Object { $_.Busy -ge 95 }).Count / [math]::Max(1, $sys.Count))
Add '--- Sistema ---'
Add ("Disco ocupado: media {0:N0}% | max {1:N0}% | amostras >= 95%: {2}%" -f $busy.Average, $busy.Maximum, $pct100)
Add ("Fila de disco: media {0:N1} | max {1:N0}" -f $q.Average, $q.Maximum)
Add ("Page Reads/s (leitura do arquivo de paginacao): media {0:N1} | max {1:N0}" -f $pr.Average, $pr.Maximum)
Add ("RAM disponivel: media {0:N0} MB | minimo {1:N0} MB" -f $av.Average, $av.Minimum)

$out | Set-Content -Path $file -Encoding UTF8
Write-Host "`nRelatorio salvo em: $file" -ForegroundColor Green
Start-Process explorer.exe "/select,`"$file`""
