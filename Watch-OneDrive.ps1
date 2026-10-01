<#
.SYNOPSIS
    Monitor continuo e leve do impacto do OneDrive. Grava uma linha por minuto
    num CSV diario, para comparar dias com e sem cada ajuste.

.DESCRIPTION
    Nao altera nada no sistema (so le contadores e registro).

    O que e medido em cada intervalo (padrao 60 s):
      - Disco: % ocupado do disco onde fica a pasta do OneDrive, amostrado a
        cada DiskInterval s (padrao 5). Grava media, maximo e a % do minuto em
        que o disco ficou >= 95% ("disco em 100%").
      - Por grupo de processo (OneDrive.Sync.Service, OneDrive, outros do
        OneDrive, Defender, indexador): RAM (working set e privada), disco
        (IO MB/s), CPU %.
      - Sistema: Page Reads/s (leitura do arquivo de paginacao), RAM livre, commit.
      - O processo de fora dessa lista que mais fez I/O no intervalo.
      - Configuracao em vigor: o que esta no IFEO (Config) e se o processo
        de sync realmente esta com a prioridade/teto (AjusteAtivo, SyncPrioIo,
        SyncWsTetoMB) - o IFEO so vale depois que o OneDrive reinicia.

    Usa contadores BRUTOS (Win32_PerfRawData_*) e calcula a media real do
    intervalo, em vez de um retrato instantaneo. Nomes de classe WMI nao sao
    traduzidos, entao funciona em Windows pt-BR (Get-Counter nao).

    O CSV usa o separador e o decimal do Windows (pt-BR: ';' e ','), entao abre
    direto no Excel. Com o arquivo aberto no Excel a gravacao falha; as linhas
    ficam em memoria e sao gravadas na proxima tentativa.

    Rode como o USUARIO (sem elevar). So um monitor por sessao.

.PARAMETER Interval
    Segundos entre linhas do CSV (30-600, padrao 60).

.PARAMETER DiskInterval
    Segundos entre amostras de disco dentro do intervalo (padrao 5).

.PARAMETER Label
    Texto livre gravado em cada linha (ex.: base, prioridade, ignore-tmp).
    A configuracao do IFEO ja e gravada sozinha na coluna Config.

.PARAMETER DurationMinutes
    Para depois de N minutos. 0 (padrao) = roda ate fechar a janela / logoff.

.PARAMETER Report
    Nao mede: le os CSVs ja gravados e mostra o resumo por dia e configuracao.

.PARAMETER FromHour / ToHour
    Com -Report: considera so linhas entre essas horas (ex.: 8 e 18 = expediente).

.EXAMPLE
    .\Watch-OneDrive.ps1 -Label base
    Mede ate fechar a janela.

.EXAMPLE
    .\Watch-OneDrive.ps1 -Report -FromHour 8 -ToHour 18
    Resumo dos dias medidos, so no horario de expediente.
#>

[CmdletBinding()]
param(
    [ValidateRange(30, 600)]
    [int]$Interval = 60,

    [ValidateRange(2, 30)]
    [int]$DiskInterval = 5,

    [string]$Label = '',

    [int]$DurationMinutes = 0,

    [switch]$Report,

    # Cria / remove o atalho na pasta Inicializar do usuario (shell:startup).
    [switch]$InstallStartup,
    [switch]$RemoveStartup,

    [ValidateRange(0, 23)]
    [int]$FromHour = 0,

    [ValidateRange(1, 24)]
    [int]$ToHour = 24,

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

$ErrorActionPreference = 'Stop'
$Culture = [Globalization.CultureInfo]::CurrentCulture

# Grupos de processos (nome sem .exe). Instancias repetidas (#1, #2) sao somadas.
$Groups = [ordered]@{
    Sync     = @('OneDrive.Sync.Service')
    Od       = @('OneDrive')
    OdOutros = @('FileCoAuth', 'Microsoft.SharePoint', 'OneDriveStandaloneUpdater', 'OneDriveSetup')
    Defender = @('MsMpEng', 'MpDefenderCoreService', 'NisSrv')
    Index    = @('SearchIndexer', 'SearchProtocolHost', 'SearchFilterHost')
}

$IfeoRoot = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options'
$CpuExpected = @{ 1 = 'Idle'; 5 = 'BelowNormal' }

# ---------------------------------------------------------------------------
# Utilitarios
# ---------------------------------------------------------------------------

function Rnd([double]$v, [int]$d = 1) { [math]::Round($v, $d) }

function N($v) {
    if ($null -eq $v -or $v -eq '') { return $null }
    [double]::Parse($v, $Culture)
}

# Leitura do estado real do processo (prioridade de I/O e teto de working set),
# que nao aparece no Get-Process. So consulta, nao altera nada.
if (-not ('OdtNative' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class OdtNative {
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
    [DllImport("kernel32.dll")]
    static extern bool CloseHandle(IntPtr h);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool GetProcessWorkingSetSizeEx(IntPtr h, out UIntPtr min, out UIntPtr max, out uint flags);
    [DllImport("ntdll.dll")]
    static extern int NtQueryInformationProcess(IntPtr h, int infoClass, ref int info, int len, out int retLen);

    // 0 VeryLow, 1 Low, 2 Normal; -1 = nao foi possivel ler
    public static int IoPriority(int pid) {
        IntPtr h = OpenProcess(0x0400, false, pid);   // PROCESS_QUERY_INFORMATION
        if (h == IntPtr.Zero) return -1;
        try {
            int v = 0, r;
            return NtQueryInformationProcess(h, 33, ref v, 4, out r) == 0 ? v : -1;   // ProcessIoPriority
        } finally { CloseHandle(h); }
    }

    // Teto rigido de working set em KB; 0 = sem teto rigido; -1 = nao foi possivel ler
    public static long HardWsMaxKB(int pid) {
        IntPtr h = OpenProcess(0x1000, false, pid);   // PROCESS_QUERY_LIMITED_INFORMATION
        if (h == IntPtr.Zero) return -1;
        try {
            UIntPtr mn, mx; uint fl;
            if (!GetProcessWorkingSetSizeEx(h, out mn, out mx, out fl)) return -1;
            return (fl & 0x4) != 0 ? (long)(mx.ToUInt64() / 1024) : 0;   // QUOTA_LIMITS_HARDWS_MAX_ENABLE
        } finally { CloseHandle(h); }
    }
}
'@
}

function Get-IfeoValues([string]$Exe) {
    $po = Join-Path $IfeoRoot "$Exe\PerfOptions"
    if (Test-Path $po) { Get-ItemProperty $po } else { $null }
}

# Resumo curto do IFEO de um exe, ex.: "cpu5-io1-ws600" ou "nada".
function Get-IfeoTag([string]$Exe) {
    $v = Get-IfeoValues $Exe
    if (-not $v) { return 'nada' }
    $parts = @()
    if ($null -ne $v.CpuPriorityClass)    { $parts += "cpu$($v.CpuPriorityClass)" }
    if ($null -ne $v.IoPriority)          { $parts += "io$($v.IoPriority)" }
    if ($null -ne $v.PagePriority)        { $parts += "pg$($v.PagePriority)" }
    if ($null -ne $v.WorkingSetLimitInKB) { $parts += "ws$([int]($v.WorkingSetLimitInKB / 1024))" }
    if ($parts) { $parts -join '-' } else { 'nada' }
}

function Get-ConfigTag {
    $s = Get-IfeoTag 'OneDrive.Sync.Service.exe'
    $o = Get-IfeoTag 'OneDrive.exe'
    $tag = if ($s -eq $o) { $s } else { "sync:$s od:$o" }
    if (Test-Path 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive') { $tag += ' +politicas' }
    $tag
}

# Letra do disco da pasta do OneDrive (cai para o disco do sistema).
function Get-OdDrive {
    $uf = Get-ChildItem 'HKCU:\Software\Microsoft\OneDrive\Accounts' -ErrorAction SilentlyContinue |
        ForEach-Object { (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).UserFolder } |
        Where-Object { $_ } | Select-Object -First 1
    if ($uf) { $uf.Substring(0, 2).ToUpper() } else { $env:SystemDrive.ToUpper() }
}

function Get-DiskRaw { @(Get-CimInstance Win32_PerfRawData_PerfDisk_PhysicalDisk -Filter "Name<>'_Total'") }
function Get-MemRaw  { Get-CimInstance Win32_PerfRawData_PerfOS_Memory }

# Um objeto por PID. Nao usa o nome da instancia como chave (OneDrive#1 muda quando processos saem).
function Get-ProcRaw {
    $t = @{}
    foreach ($p in Get-CimInstance Win32_PerfRawData_PerfProc_Process) {
        if ($p.Name -eq '_Total' -or $p.Name -eq 'Idle') { continue }
        $t[[int]$p.IDProcess] = $p
    }
    $t
}

# Diferenca de um contador UInt32 (latencia e transferencias de disco), que da a volta em 2^32.
# O tempo de latencia acumulado (10 MHz) estoura a cada poucos minutos de I/O.
function Get-Delta32($Old, $New) {
    $d = [double]$New - [double]$Old
    if ($d -lt 0) { $d += 4294967296.0 }
    $d
}

# % ocupado e fila media de cada disco entre duas leituras brutas.
# Nao usa o _Total: ele e a MEDIA dos discos, e esconde um disco em 100% ao lado de um ocioso.
function Get-DiskBusy($a, $b) {
    foreach ($x in $b) {
        $o = $a | Where-Object { $_.Name -eq $x.Name } | Select-Object -First 1
        if (-not $o) { continue }
        $dt = [double]$x.PercentIdleTime_Base - [double]$o.PercentIdleTime_Base
        if ($dt -le 0) { continue }
        $idle = ([double]$x.PercentIdleTime - [double]$o.PercentIdleTime) / $dt
        # Em SSD, "% ocupado" = havia pelo menos 1 operacao pendente; IOPS e latencia mostram se saturou de fato.
        $secs = ([double]$x.Timestamp_PerfTime - [double]$o.Timestamp_PerfTime) / [double]$x.Frequency_PerfTime
        $n    = Get-Delta32 $o.DiskTransfersPersec $x.DiskTransfersPersec
        $nb   = Get-Delta32 $o.AvgDisksecPerTransfer_Base $x.AvgDisksecPerTransfer_Base
        $lat  = if ($nb -gt 0) { (Get-Delta32 $o.AvgDisksecPerTransfer $x.AvgDisksecPerTransfer) / [double]$x.Frequency_PerfTime / $nb * 1000 } else { 0 }
        [pscustomobject]@{
            Name  = $x.Name
            Busy  = [math]::Min(100, [math]::Max(0, 100 * (1 - $idle)))
            Queue = ([double]$x.AvgDiskQueueLength - [double]$o.AvgDiskQueueLength) / $dt
            Iops  = if ($secs -gt 0) { [math]::Max(0, $n) / $secs } else { 0 }
            LatMs = [math]::Max(0, $lat)
        }
    }
}

# ---------------------------------------------------------------------------
# Relatorio
# ---------------------------------------------------------------------------

function Get-Stat($values, [string]$Kind) {
    $v = @($values | Where-Object { $null -ne $_ } | Sort-Object)
    if ($v.Count -eq 0) { return $null }
    switch ($Kind) {
        'avg' { Rnd ($v | Measure-Object -Average).Average }
        'p95' { Rnd $v[[math]::Max(0, [int][math]::Ceiling(0.95 * $v.Count) - 1)] }
        'max' { Rnd $v[-1] }
    }
}

function Col($rows, [string]$Name) { foreach ($r in $rows) { N $r.$Name } }

function Get-Mode($rows, [string]$Name) {
    ($rows | Group-Object $Name | Sort-Object Count -Descending | Select-Object -First 1).Name
}

function Show-Report {
    $files = @(Get-ChildItem $OutDir -Filter 'onedrive-watch-*.csv' -ErrorAction SilentlyContinue | Sort-Object Name)
    if (-not $files) { Write-Host "Nenhum CSV do monitor em $OutDir"; return }

    $rows = foreach ($f in $files) { Import-Csv $f.FullName -UseCulture }
    $rows = @($rows | Where-Object {
        $h = [int]$_.Hora.Substring(11, 2); $h -ge $FromHour -and $h -lt $ToHour })
    if (-not $rows) { Write-Host 'Nenhuma linha no horario pedido.'; return }

    $groups = $rows | Group-Object { $_.Hora.Substring(0, 10) }, Config, Rotulo

    $summary = foreach ($g in $groups) {
        $r = $g.Group
        $totRam = foreach ($x in $r) { (N $x.SyncRamMB) + (N $x.OdRamMB) + (N $x.OdOutrosRamMB) }
        [pscustomobject]@{
            Dia           = $r[0].Hora.Substring(0, 10)
            Config        = $r[0].Config
            Rotulo        = $r[0].Rotulo
            Ativo         = Get-Mode $r 'AjusteAtivo'
            Min           = $r.Count
            'Disco%'      = Get-Stat (Col $r 'DiscoOcupMed') 'avg'
            'Em100%'      = Get-Stat (Col $r 'DiscoPct95') 'avg'
            MinSat        = @($r | Where-Object { (N $_.DiscoPct95) -ge 50 }).Count
            Iops          = Get-Stat (Col $r 'DiscoIops') 'avg'
            LatMs         = Get-Stat (Col $r 'DiscoLatMs') 'avg'
            LatP95        = Get-Stat (Col $r 'DiscoLatMs') 'p95'
            PgRd          = Get-Stat (Col $r 'PageReadsS') 'avg'
            PgRdP95       = Get-Stat (Col $r 'PageReadsS') 'p95'
            LivreMB       = Get-Stat (Col $r 'RamLivreMB') 'avg'
            SyncPrivMB    = Get-Stat (Col $r 'SyncPrivMB') 'avg'
            SyncMB        = Get-Stat (Col $r 'SyncRamMB') 'avg'
            SyncP95       = Get-Stat (Col $r 'SyncRamMB') 'p95'
            OdMB          = Get-Stat (Col $r 'OdRamMB') 'avg'
            TotOdMB       = Get-Stat $totRam 'avg'
            TetoWS        = Get-Mode $r 'SyncWsTetoMB'
            SyncIO        = Get-Stat (Col $r 'SyncIoMBs') 'avg'
            OdIO          = Get-Stat (Col $r 'OdIoMBs') 'avg'
            DefIO         = Get-Stat (Col $r 'DefenderIoMBs') 'avg'
            IdxIO         = Get-Stat (Col $r 'IndexIoMBs') 'avg'
        }
    }

    # Quem fazia I/O nos minutos em que o disco passou pelo menos metade do tempo em 100%.
    $sat = foreach ($g in ($rows | Where-Object { (N $_.DiscoPct95) -ge 50 } |
                           Group-Object { $_.Hora.Substring(0, 10) }, Config)) {
        $r = $g.Group
        [pscustomobject]@{
            Dia      = $r[0].Hora.Substring(0, 10)
            Config   = $r[0].Config
            MinSat   = $r.Count
            SyncIO   = Get-Stat (Col $r 'SyncIoMBs') 'avg'
            OdIO     = Get-Stat (Col $r 'OdIoMBs') 'avg'
            DefIO    = Get-Stat (Col $r 'DefenderIoMBs') 'avg'
            IdxIO    = Get-Stat (Col $r 'IndexIoMBs') 'avg'
            SyncOps  = Get-Stat (Col $r 'SyncOpsS') 'avg'
            OdOps    = Get-Stat (Col $r 'OdOpsS') 'avg'
            DefOps   = Get-Stat (Col $r 'DefenderOpsS') 'avg'
            IdxOps   = Get-Stat (Col $r 'IndexOpsS') 'avg'
            PgRd     = Get-Stat (Col $r 'PageReadsS') 'avg'
            TopFora  = Get-Mode $r 'TopIoProc'
            TopIO    = Get-Stat (Col $r 'TopIoMBs') 'avg'
            TopOps   = Get-Mode $r 'TopOpsProc'
        }
    }

    # Memoria: quem aparece no top 5 de memoria privada, por dia. Processo fora do top 5
    # num minuto conta 0 naquele minuto, entao a media e um piso (aproximada).
    $mem = foreach ($g in ($rows | Where-Object { $_.TopMemMB } | Group-Object { $_.Hora.Substring(0, 10) })) {
        $n = $g.Count
        $g.Group | ForEach-Object { $_.TopMemMB -split ' / ' } | ForEach-Object {
            $i = $_.LastIndexOf(' ')
            [pscustomobject]@{ Proc = $_.Substring(0, $i); MB = [double]$_.Substring($i + 1) }
        } | Group-Object Proc | ForEach-Object {
            [pscustomobject]@{
                Dia        = $g.Name
                Processo   = $_.Name
                'MB medio' = [int](($_.Group | Measure-Object MB -Sum).Sum / $n)
                'MB max'   = [int]($_.Group | Measure-Object MB -Maximum).Maximum
                'No top5%' = [int](100 * $_.Count / $n)
            }
        } | Sort-Object 'MB medio' -Descending | Select-Object -First 8
    }

    Write-Host ("`n== Resumo por dia e configuracao (horas {0}-{1}) ==" -f $FromHour, $ToHour) -ForegroundColor Cyan
    Write-Host 'Em100% = % do tempo com disco >= 95% | MinSat = minutos com disco em 100% por metade do minuto ou mais'
    Write-Host 'RAM em MB: SyncPrivMB = memoria privada (custo real); *MB = working set (o Windows reduz sozinho) | IO em MB/s | PgRd = Page Reads/s'
    # Duas tabelas para caber no console; o CSV do resumo leva todas as colunas.
    Write-Host 'Iops/LatMs = operacoes e latencia media do disco (em SSD, latencia alta = saturado de verdade)'
    Write-Host ($summary | Format-Table Dia, Config, Rotulo, Ativo, Min, 'Disco%', 'Em100%', MinSat, Iops, LatMs, LatP95, PgRd, PgRdP95, LivreMB -AutoSize |
        Out-String -Width 400).TrimEnd()
    Write-Host ($summary | Format-Table Dia, Config, Rotulo, SyncPrivMB, SyncMB, SyncP95, OdMB, TotOdMB, TetoWS, SyncIO, OdIO, DefIO, IdxIO -AutoSize |
        Out-String -Width 400).TrimEnd()

    Write-Host "`n== Nos minutos saturados: quem fazia I/O ==" -ForegroundColor Cyan
    if ($sat) { Write-Host ($sat | Format-Table * -AutoSize | Out-String -Width 400).TrimEnd() }
    else      { Write-Host '  (nenhum minuto saturado)' }
    Write-Host 'Ops = operacoes de I/O por segundo, incluindo abrir/listar/ler atributos (o MB/s nao ve isso)'

    Write-Host "`n== Memoria: top processos por memoria privada (MB) ==" -ForegroundColor Cyan
    if ($mem) { Write-Host ($mem | Format-Table * -AutoSize | Out-String -Width 400).TrimEnd() }
    else      { Write-Host '  (sem dados - CSV de versao anterior)' }

    $csv = Join-Path $OutDir ("onedrive-resumo-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmm'))
    $summary | Export-Csv -Path $csv -NoTypeInformation -UseCulture -Encoding UTF8
    Write-Host "`nResumo salvo em: $csv" -ForegroundColor Green
}

if ($Report) { Show-Report; return }

# Atalho na pasta Inicializar do USUARIO (nao e tarefa agendada nem registro): o monitor
# abre minimizado a cada login. Para parar de monitorar: -RemoveStartup.
$StartupLnk = Join-Path ([Environment]::GetFolderPath('Startup')) 'Watch-OneDrive.lnk'
if ($RemoveStartup) {
    if (Test-Path $StartupLnk) { Remove-Item $StartupLnk; Write-Host "Atalho removido: $StartupLnk" -ForegroundColor Green }
    else { Write-Host 'Nao havia atalho na pasta Inicializar.' }
    return
}
if ($InstallStartup) {
    $lnkArgs = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Minimized -File "{0}"' -f $PSCommandPath
    if ($Label) { $lnkArgs += ' -Label "{0}"' -f $Label }
    $sh  = New-Object -ComObject WScript.Shell
    $lnk = $sh.CreateShortcut($StartupLnk)
    $lnk.TargetPath       = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $lnk.Arguments        = $lnkArgs
    $lnk.WorkingDirectory = $PSScriptRoot
    $lnk.WindowStyle      = 7   # minimizado
    $lnk.Description      = 'Monitor do OneDrive (onedrive-throttle)'
    $lnk.Save()
    Write-Host "Atalho criado: $StartupLnk" -ForegroundColor Green
    Write-Host "  powershell.exe $lnkArgs"
    Write-Host 'O monitor abre minimizado no proximo login. Para remover: -RemoveStartup'
    return
}

# ---------------------------------------------------------------------------
# Monitor
# ---------------------------------------------------------------------------

$created = $false
$mutex = New-Object System.Threading.Mutex($true, 'Local\onedrive-throttle-watch', [ref]$created)
if (-not $created) { Write-Warning 'Ja existe um Watch-OneDrive rodando nesta sessao.'; return }

try {
    if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
    try { (Get-Process -Id $PID).PriorityClass = 'BelowNormal' } catch {}   # o proprio monitor nao disputa CPU

    $ncpu    = [int](Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors
    $odDrive = Get-OdDrive
    $known   = @{}
    foreach ($k in $Groups.Keys) { foreach ($n in $Groups[$k]) { $known[$n] = $k } }

    Write-Host ("Monitor iniciado: linha a cada {0} s, disco {1} a cada {2} s. CSV em {3}" -f $Interval, $odDrive, $DiskInterval, $OutDir) -ForegroundColor Yellow
    Write-Host 'Pode minimizar esta janela. Ctrl+C para parar (o que ja foi gravado fica).'

    $prevProc = Get-ProcRaw
    $prevMem  = Get-MemRaw
    $prevDisk = Get-DiskRaw
    $win      = New-Object System.Collections.Generic.List[object]
    $pending  = New-Object System.Collections.Generic.List[object]
    $stopAt   = if ($DurationMinutes -gt 0) { (Get-Date).AddMinutes($DurationMinutes) } else { [datetime]::MaxValue }
    $next     = (Get-Date).AddSeconds($Interval)

    while ((Get-Date) -lt $stopAt) {
        Start-Sleep -Seconds $DiskInterval

        # --- Disco, a cada DiskInterval ---
        $curDisk = Get-DiskRaw
        $busy = @(Get-DiskBusy $prevDisk $curDisk)
        $prevDisk = $curDisk
        $od = $busy | Where-Object { ($_.Name -split ' ') -contains $odDrive } | Select-Object -First 1
        if ($busy) {
            $win.Add([pscustomobject]@{
                Od    = if ($od) { $od.Busy } else { $null }
                Queue = if ($od) { $od.Queue } else { $null }
                Iops  = if ($od) { $od.Iops } else { $null }
                LatMs = if ($od) { $od.LatMs } else { $null }
                Any   = ($busy | Measure-Object Busy -Maximum).Maximum
            })
        }

        if ((Get-Date) -lt $next) { continue }
        $now  = Get-Date
        $next = $now.AddSeconds($Interval)

        # --- Processos e memoria, a cada Interval ---
        $curProc = Get-ProcRaw
        $curMem  = Get-MemRaw
        $secs = ([double]$curMem.Timestamp_PerfTime - [double]$prevMem.Timestamp_PerfTime) / [double]$curMem.Frequency_PerfTime

        # Depois de suspender/hibernar o intervalo fica enorme e a media nao significa nada: descarta.
        if ($secs -le 0 -or $secs -gt 2 * $Interval) {
            $prevProc = $curProc; $prevMem = $curMem; $win.Clear()
            continue
        }

        # Soma por nome (sem #n). Processo que nasceu no intervalo conta desde zero.
        $agg = @{}
        foreach ($procId in $curProc.Keys) {
            $x    = $curProc[$procId]
            $name = $x.Name -replace '#\d+$', ''
            $o    = $prevProc[$procId]
            if ($o -and ($o.Name -replace '#\d+$', '') -ne $name) { $o = $null }   # PID reaproveitado
            $dIo  = [double]$x.IODataBytesPersec - $(if ($o) { [double]$o.IODataBytesPersec } else { 0 })
            $dCpu = [double]$x.PercentProcessorTime - $(if ($o) { [double]$o.PercentProcessorTime } else { 0 })
            $dPf  = [double]$x.PageFaultsPersec - $(if ($o) { [double]$o.PageFaultsPersec } else { 0 })
            # Operacoes de I/O (dados + outras: abrir, listar pasta, ler atributos). IODataBytes nao ve
            # varredura de pastas e metadados, que e boa parte do trabalho do OneDrive/Defender/indexador.
            $dOps = [double]$x.IODataOperationsPersec + [double]$x.IOOtherOperationsPersec -
                    $(if ($o) { [double]$o.IODataOperationsPersec + [double]$o.IOOtherOperationsPersec } else { 0 })
            if (-not $agg.ContainsKey($name)) {
                $agg[$name] = [pscustomobject]@{ Io = 0.0; Ops = 0.0; Cpu = 0.0; Ws = 0.0; Priv = 0.0; Pf = 0.0 }
            }
            $a = $agg[$name]
            $a.Io   += [math]::Max(0, $dIo) / $secs
            $a.Ops  += [math]::Max(0, $dOps) / $secs
            $a.Cpu  += [math]::Max(0, $dCpu) / ($secs * 1e7) / $ncpu * 100
            $a.Pf   += [math]::Max(0, $dPf) / $secs
            $a.Ws   += [double]$x.WorkingSet
            $a.Priv += [double]$x.PrivateBytes
        }

        $grp = @{}
        foreach ($k in $Groups.Keys) { $grp[$k] = [pscustomobject]@{ Io = 0.0; Ops = 0.0; Cpu = 0.0; Ws = 0.0; Priv = 0.0; Pf = 0.0; Found = $false } }
        $top = $null; $topOps = $null
        foreach ($name in $agg.Keys) {
            $a = $agg[$name]
            if ($known.ContainsKey($name)) {
                $g = $grp[$known[$name]]
                $g.Io += $a.Io; $g.Ops += $a.Ops; $g.Cpu += $a.Cpu; $g.Ws += $a.Ws; $g.Priv += $a.Priv; $g.Pf += $a.Pf; $g.Found = $true
            } else {
                if (-not $top -or $a.Io -gt $agg[$top].Io) { $top = $name }
                if (-not $topOps -or $a.Ops -gt $agg[$topOps].Ops) { $topOps = $name }
            }
        }

        # Top 5 de memoria privada (todos os processos): quem empurra a maquina para o arquivo de paginacao.
        $topMem = ($agg.GetEnumerator() | Sort-Object { $_.Value.Priv } -Descending | Select-Object -First 5 |
            ForEach-Object { '{0} {1}' -f $_.Key, [int]($_.Value.Priv / 1MB) }) -join ' / '

        # Estado real do motor de sync (o IFEO so vale apos reiniciar o OneDrive).
        $sp = Get-Process -Name 'OneDrive.Sync.Service' -ErrorAction SilentlyContinue | Select-Object -First 1
        $syncCpuPrio = $null; $syncIo = $null; $syncWs = $null; $ativo = 'n/a'
        if ($sp) {
            try { $syncCpuPrio = [string]$sp.PriorityClass } catch {}
            $syncIo = [OdtNative]::IoPriority($sp.Id)
            $kb = [OdtNative]::HardWsMaxKB($sp.Id)
            $syncWs = if ($kb -ge 0) { [int]($kb / 1024) } else { $null }
            $v = Get-IfeoValues 'OneDrive.Sync.Service.exe'
            if ($v -and ($null -ne $v.CpuPriorityClass -or $null -ne $v.IoPriority)) {
                $ok = $true
                if ($null -ne $v.CpuPriorityClass -and $CpuExpected[[int]$v.CpuPriorityClass] -ne $syncCpuPrio) { $ok = $false }
                if ($null -ne $v.IoPriority -and $syncIo -ge 0 -and $syncIo -ne [int]$v.IoPriority) { $ok = $false }
                $ativo = if ($ok) { 'sim' } else { 'nao' }
            }
        }

        $odW  = @($win | Where-Object { $null -ne $_.Od })
        $dAvg = if ($odW) { ($odW | Measure-Object Od -Average).Average } else { $null }
        $dMax = if ($odW) { ($odW | Measure-Object Od -Maximum).Maximum } else { $null }
        $dP95 = if ($odW) { 100 * @($odW | Where-Object { $_.Od -ge 95 }).Count / $odW.Count } else { $null }
        $dQ   = if ($odW) { ($odW | Measure-Object Queue -Average).Average } else { $null }
        $aP95 = if ($win.Count) { 100 * @($win | Where-Object { $_.Any -ge 95 }).Count / $win.Count } else { $null }
        $dIops = if ($odW) { ($odW | Measure-Object Iops -Average).Average } else { $null }
        # Latencia media ponderada pelas operacoes de cada janela.
        $sumN  = ($odW | Measure-Object Iops -Sum).Sum
        $dLat  = if ($sumN -gt 0) { (($odW | ForEach-Object { $_.Iops * $_.LatMs }) | Measure-Object -Sum).Sum / $sumN } else { $null }

        $pr = ([double]$curMem.PageReadsPersec - [double]$prevMem.PageReadsPersec) / $secs
        $S = $grp['Sync']; $O = $grp['Od']; $X = $grp['OdOutros']; $D = $grp['Defender']; $I = $grp['Index']

        $row = [pscustomobject][ordered]@{
            Hora            = $now.ToString('yyyy-MM-dd HH:mm:ss')
            Rotulo          = $Label
            Config          = Get-ConfigTag
            AjusteAtivo     = $ativo
            SyncPid         = if ($sp) { $sp.Id } else { $null }
            SyncPrioCpu     = $syncCpuPrio
            SyncPrioIo      = $syncIo
            SyncWsTetoMB    = $syncWs
            DiscoOcupMed    = if ($null -ne $dAvg) { Rnd $dAvg } else { $null }
            DiscoOcupMax    = if ($null -ne $dMax) { Rnd $dMax } else { $null }
            DiscoPct95      = if ($null -ne $dP95) { Rnd $dP95 } else { $null }
            DiscoFila       = if ($null -ne $dQ)   { Rnd $dQ 2 } else { $null }
            DiscoIops       = if ($null -ne $dIops) { Rnd $dIops 0 } else { $null }
            DiscoLatMs      = if ($null -ne $dLat)  { Rnd $dLat 2 } else { $null }
            QualquerDiscoPct95 = if ($null -ne $aP95) { Rnd $aP95 } else { $null }
            PageReadsS      = Rnd $pr
            RamLivreMB      = [int]$curMem.AvailableMBytes
            CommitMB        = [int]([double]$curMem.CommittedBytes / 1MB)
            SyncRamMB       = if ($S.Found) { [int]($S.Ws / 1MB) } else { $null }
            SyncPrivMB      = if ($S.Found) { [int]($S.Priv / 1MB) } else { $null }
            SyncIoMBs       = Rnd ($S.Io / 1MB) 2
            SyncOpsS        = Rnd $S.Ops 0
            SyncCpuPct      = Rnd $S.Cpu
            SyncFaltasS     = Rnd $S.Pf 0
            OdRamMB         = if ($O.Found) { [int]($O.Ws / 1MB) } else { $null }
            OdPrivMB        = if ($O.Found) { [int]($O.Priv / 1MB) } else { $null }
            OdIoMBs         = Rnd ($O.Io / 1MB) 2
            OdOpsS          = Rnd $O.Ops 0
            OdCpuPct        = Rnd $O.Cpu
            OdOutrosRamMB   = [int]($X.Ws / 1MB)
            OdOutrosIoMBs   = Rnd ($X.Io / 1MB) 2
            DefenderIoMBs   = Rnd ($D.Io / 1MB) 2
            DefenderOpsS    = Rnd $D.Ops 0
            DefenderCpuPct  = Rnd $D.Cpu
            IndexIoMBs      = Rnd ($I.Io / 1MB) 2
            IndexOpsS       = Rnd $I.Ops 0
            IndexCpuPct     = Rnd $I.Cpu
            TopIoProc       = $top
            TopIoMBs        = if ($top) { Rnd ($agg[$top].Io / 1MB) 2 } else { $null }
            TopOpsProc      = $topOps
            TopOpsS         = if ($topOps) { Rnd $agg[$topOps].Ops 0 } else { $null }
            TopMemMB        = $topMem
        }

        $prevProc = $curProc; $prevMem = $curMem; $win.Clear()

        # Grava (um arquivo por dia). Se o CSV estiver aberto no Excel, tenta de novo no proximo minuto.
        $pending.Add($row)
        try {
            foreach ($day in @($pending | Group-Object { $_.Hora.Substring(0, 10).Replace('-', '') })) {
                $file = Join-Path $OutDir ("onedrive-watch-{0}-{1}.csv" -f $env:COMPUTERNAME, $day.Name)
                # CSV do dia gravado por uma versao com outras colunas: Export-Csv -Append recusaria
                # para sempre. Renomeia o antigo (o -Report continua lendo os dois).
                if (Test-Path $file) {
                    $head = (Get-Content $file -TotalCount 1) -replace '"', ''
                    if ($head -ne ($day.Group[0].PSObject.Properties.Name -join $Culture.TextInfo.ListSeparator)) {
                        Rename-Item $file ("onedrive-watch-{0}-{1}-anterior-{2}.csv" -f $env:COMPUTERNAME, $day.Name, (Get-Date -Format 'HHmmss'))
                    }
                }
                $day.Group | Export-Csv -Path $file -Append -NoTypeInformation -UseCulture -Encoding UTF8
            }
            $pending.Clear()
        } catch {
            Write-Warning ("Nao gravei agora ({0} linha(s) pendentes - CSV aberto no Excel?): {1}" -f $pending.Count, $_.Exception.Message)
        }

        Write-Host ("{0}  disco {1,3}% (em 100%: {2,3}% do minuto, {3} IOPS, {4} ms)  sync {5,5} MB  od {6,4} MB  pgrd {7,5}/s  fora: {8} {9} MB/s  [{10}]" -f
            $now.ToString('HH:mm'), $row.DiscoOcupMed, $row.DiscoPct95, $row.DiscoIops, $row.DiscoLatMs, $row.SyncRamMB, $row.OdRamMB,
            $row.PageReadsS, $row.TopIoProc, $row.TopIoMBs, $row.Config)
    }
} finally {
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
