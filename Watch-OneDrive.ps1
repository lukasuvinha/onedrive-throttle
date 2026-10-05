<#
.SYNOPSIS
    Monitor continuo e leve do impacto do OneDrive. Grava uma linha por minuto
    num CSV diario, para comparar dias com e sem cada ajuste.

.DESCRIPTION
    Nao altera nada no sistema (so le contadores e registro).

    O que e medido em cada intervalo (padrao 60 s):
      - TRAVAMENTOS (o problema que o usuario sente): a cada DiskInterval s,
        quantas janelas estao "Nao respondendo" (Get-Process com janela
        principal + IsHungAppWindow, o mesmo criterio do titulo "(Nao
        respondendo)" do Windows). Por minuto: em quantas amostras houve
        janela travada (TravaAmostras de Amostras), o maximo de janelas
        travadas juntas e quais processos travaram (TravaProcs).
        Nao usa Process.Responding: ele acusa app UWP suspenso (ex.:
        Configuracoes minimizado) e pode bloquear ate 5 s.
      - CPU total da maquina: media e pico das amostras (CpuTotalPct/Max).
      - Disco onde fica a pasta do OneDrive, amostrado a cada DiskInterval s
        (padrao 5): LATENCIA media (metrica principal: em SSD e ela que mostra
        saturacao), IOPS, % ocupado e a % do minuto em que ficou >= 95%.
        Em SSD, % ocupado 100% so quer dizer "sempre havia 1 operacao pendente".
      - Por grupo de processo (OneDrive.Sync.Service, OneDrive, outros do
        OneDrive, antivirus [ESET e Defender], indexador): RAM (working set e
        privada), disco (IO MB/s e operacoes/s), CPU %.
      - Sistema: Page Reads/s, RAM livre, commit e uso do arquivo de paginacao.
        Page Reads/s NAO e so arquivo de paginacao: conta toda leitura de disco
        feita por falta de pagina, inclusive arquivos que nao estavam no cache
        do Windows e arquivos mapeados em memoria. Para falta de RAM, use o uso
        do arquivo de paginacao (PaginacaoUsoPct/MB) junto com a RAM livre.
      - O processo de fora dessa lista que mais fez I/O no intervalo.
      - Inicio do motor de sync (SyncInicio): o -Report separa os primeiros
        60 min apos cada inicio do OneDrive (sincronizacao inicial) do resto.
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

.PARAMETER SlowLatMs
    Com -Report: latencia media (ms) a partir da qual o minuto conta como
    "lento" (padrao 20). SSD SATA saudavel fica abaixo de ~5 ms.

.PARAMETER StartMinutes
    Com -Report: quantos minutos apos cada inicio do OneDrive contam como
    fase "inicio" (padrao 60).

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

    [ValidateRange(1, 1000)]
    [double]$SlowLatMs = 20,

    [ValidateRange(5, 600)]
    [int]$StartMinutes = 60,

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

$ErrorActionPreference = 'Stop'
$Culture = [Globalization.CultureInfo]::CurrentCulture

# Grupos de processos (nome sem .exe). Instancias repetidas (#1, #2) sao somadas.
$Groups = [ordered]@{
    Sync     = @('OneDrive.Sync.Service')
    Od       = @('OneDrive')
    OdOutros = @('FileCoAuth', 'Microsoft.SharePoint', 'OneDriveStandaloneUpdater', 'OneDriveSetup')
    # Antivirus: ESET (o da empresa) e Defender (pode estar ativo em paralelo ou em outras maquinas).
    Antivirus = @('ekrn', 'egui', 'MsMpEng', 'MpDefenderCoreService', 'NisSrv')
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
    [DllImport("user32.dll")]
    static extern bool IsHungAppWindow(IntPtr hWnd);

    // Mesmo criterio do "(Nao respondendo)" do Windows: a janela nao processa mensagens ha ~5 s.
    // Nao bloqueia (Process.Responding pode esperar ate 5 s) e nao acusa app UWP suspenso.
    public static bool IsHung(IntPtr hWnd) { return hWnd != IntPtr.Zero && IsHungAppWindow(hWnd); }

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

# CPU total da maquina (_Total de todos os nucleos). PercentProcessorTime bruto conta o tempo
# OCIOSO em 100 ns (contador invertido): ocupado = 1 - ocioso / tempo decorrido.
function Get-CpuRaw { Get-CimInstance Win32_PerfRawData_PerfOS_Processor -Filter "Name='_Total'" }
function Get-CpuBusy($a, $b) {
    $dt = [double]$b.Timestamp_Sys100NS - [double]$a.Timestamp_Sys100NS
    if ($dt -le 0) { return $null }
    [math]::Min(100.0, [math]::Max(0.0, 100 * (1 - ([double]$b.PercentProcessorTime - [double]$a.PercentProcessorTime) / $dt)))
}

# Janelas "Nao respondendo" agora: processos com janela principal (Get-Process / MainWindowHandle)
# cuja janela o Windows considera travada. Devolve um nome por janela travada.
function Get-HungWindows {
    foreach ($p in Get-Process) {
        if ($p.Id -eq $PID) { continue }
        # dwm: quando uma janela trava, o Windows cria uma "janela fantasma" (a copia esbranquicada
        # com "(Nao respondendo)") que pertence ao dwm.exe. Contar o dwm duplicaria cada travamento.
        if ($p.ProcessName -eq 'dwm') { continue }
        try { $h = $p.MainWindowHandle } catch { continue }
        if ($h -ne [IntPtr]::Zero -and [OdtNative]::IsHung($h)) { $p.ProcessName }
    }
}

# Uso do arquivo de paginacao (_Total). PercentUsage bruto = paginas de 4 KB em uso;
# _Base = tamanho do arquivo em paginas. Arquivo de paginacao cheio + pouca RAM livre =
# falta de RAM de verdade (Page Reads/s sozinho nao diz isso).
function Get-PagingFile {
    $p = Get-CimInstance Win32_PerfRawData_PerfOS_PagingFile -Filter "Name='_Total'" -ErrorAction SilentlyContinue
    if (-not $p -or -not [double]$p.PercentUsage_Base) { return $null }
    [pscustomobject]@{
        Pct = 100 * [double]$p.PercentUsage / [double]$p.PercentUsage_Base
        MB  = [double]$p.PercentUsage * 4KB / 1MB
    }
}

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
    if ($d -lt 0) {
        # So e volta do contador de 32 bits se o valor antigo estava perto do teto;
        # caso contrario (reinicio do contador, troca de instancia) descarta a amostra.
        if ([double]$Old -gt 3.0e9) { $d += 4294967296.0 } else { $d = 0.0 }
    }
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
            Busy  = [math]::Min(100.0, [math]::Max(0.0, [double](100 * (1 - $idle))))
            Queue = ([double]$x.AvgDiskQueueLength - [double]$o.AvgDiskQueueLength) / $dt
            Iops  = if ($secs -gt 0) { [math]::Max(0.0, [double]$n) / $secs } else { 0 }
            LatMs = [math]::Max(0.0, [double]$lat)
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

# Valores de uma coluna. Aceita nomes alternativos para ler CSVs de versoes anteriores
# (ex.: AntivirusOpsS, antes DefenderOpsS): usa o primeiro que existir na linha.
function Col($rows, [string[]]$Names) {
    foreach ($r in $rows) {
        $v = $null
        foreach ($n in $Names) { if ($null -ne $r.$n -and $r.$n -ne '') { $v = $r.$n; break } }
        N $v
    }
}

function Get-Mode($rows, [string]$Name) {
    ($rows | Group-Object $Name | Sort-Object Count -Descending | Select-Object -First 1).Name
}

# Fase de cada minuto em relacao ao inicio do motor de sync:
#   inicio   = primeiros StartMinutes minutos (sincronizacao inicial: varre tudo apos o login/restart)
#   resto    = depois disso (o efeito do trabalho do usuario no dia)
#   sem sync = OneDrive.Sync.Service nao estava rodando
# Usa SyncInicio (hora real de inicio do processo). CSVs antigos nao tem essa coluna: o inicio e
# estimado pela primeira aparicao do PID nos CSVs, o que erra se o monitor abriu depois do OneDrive.
function Add-Phase($rows) {
    $fmt = 'yyyy-MM-dd HH:mm:ss'
    $firstSeen = @{}
    foreach ($r in ($rows | Sort-Object Hora)) {
        if ($r.SyncPid -and -not $firstSeen.ContainsKey($r.SyncPid)) {
            $firstSeen[$r.SyncPid] = [datetime]::ParseExact($r.Hora, $fmt, $null)
        }
    }
    $estimated = 0
    foreach ($r in $rows) {
        $fase = 'sem sync'
        if ($r.SyncPid) {
            if ($r.SyncInicio) { $start = [datetime]::ParseExact($r.SyncInicio, $fmt, $null) }
            else { $start = $firstSeen[$r.SyncPid]; $estimated++ }
            $mins = ([datetime]::ParseExact($r.Hora, $fmt, $null) - $start).TotalMinutes
            $fase = if ($mins -lt $StartMinutes) { 'inicio' } else { 'resto' }
        }
        $r | Add-Member -NotePropertyName Fase -NotePropertyValue $fase -Force
    }
    $estimated
}

function Get-PhaseOrder([string]$Fase) { switch ($Fase) { 'inicio' { 0 } 'resto' { 1 } default { 2 } } }

function Show-Report {
    $files = @(Get-ChildItem $OutDir -Filter 'onedrive-watch-*.csv' -ErrorAction SilentlyContinue | Sort-Object Name)
    if (-not $files) { Write-Host "Nenhum CSV do monitor em $OutDir"; return }

    $rows = @(foreach ($f in $files) { Import-Csv $f.FullName -UseCulture })
    # Fase antes do filtro de horario: o inicio do sync pode ter sido antes de FromHour.
    $estimated = Add-Phase $rows
    $rows = @($rows | Where-Object {
        $h = [int]$_.Hora.Substring(11, 2); $h -ge $FromHour -and $h -lt $ToHour })
    if (-not $rows) { Write-Host 'Nenhuma linha no horario pedido.'; return }

    $slow = { (N $_.DiscoLatMs) -ge $SlowLatMs }
    # Linhas que mediram travamento (CSVs antigos nao tem a coluna) e as que tiveram trava.
    $measured = { $null -ne $_.TravaAmostras -and $_.TravaAmostras -ne '' }
    $hungRow  = { (N $_.TravaAmostras) -gt 0 }

    $groups = $rows | Group-Object { $_.Hora.Substring(0, 10) }, Config, Rotulo, Fase
    $summary = foreach ($g in $groups) {
        $r = $g.Group
        $totRam = foreach ($x in $r) { (N $x.SyncRamMB) + (N $x.OdRamMB) + (N $x.OdOutrosRamMB) }
        $rm = @($r | Where-Object $measured)
        $amostras = ($rm | ForEach-Object { N $_.Amostras } | Measure-Object -Sum).Sum
        $travadas = ($rm | ForEach-Object { N $_.TravaAmostras } | Measure-Object -Sum).Sum
        [pscustomobject]@{
            Dia        = $r[0].Hora.Substring(0, 10)
            Config     = $r[0].Config
            Rotulo     = $r[0].Rotulo
            Fase       = $r[0].Fase
            Ativo      = Get-Mode $r 'AjusteAtivo'
            Min        = $r.Count
            # O que o usuario sente: minutos com alguma janela "Nao respondendo".
            MinTrava   = if ($rm) { @($rm | Where-Object $hungRow).Count } else { $null }
            'Trava%'   = if ($amostras) { Rnd (100 * $travadas / $amostras) } else { $null }
            CpuMed     = Get-Stat (Col $r 'CpuTotalPct') 'avg'
            CpuP95     = Get-Stat (Col $r 'CpuTotalPct') 'p95'
            # Metrica principal de disco: latencia.
            LatMs      = Get-Stat (Col $r 'DiscoLatMs') 'avg'
            LatP95     = Get-Stat (Col $r 'DiscoLatMs') 'p95'
            MinLentos  = @($r | Where-Object $slow).Count
            Iops       = Get-Stat (Col $r 'DiscoIops') 'avg'
            'Ocup%'    = Get-Stat (Col $r 'DiscoOcupMed') 'avg'
            'Em100%'   = Get-Stat (Col $r 'DiscoPct95') 'avg'
            PgRd       = Get-Stat (Col $r 'PageReadsS') 'avg'
            'PagUso%'  = Get-Stat (Col $r 'PaginacaoUsoPct') 'avg'
            PagUsoMB   = Get-Stat (Col $r 'PaginacaoUsoMB') 'max'
            LivreMB    = Get-Stat (Col $r 'RamLivreMB') 'avg'
            SyncPrivMB = Get-Stat (Col $r 'SyncPrivMB') 'avg'
            SyncMB     = Get-Stat (Col $r 'SyncRamMB') 'avg'
            OdMB       = Get-Stat (Col $r 'OdRamMB') 'avg'
            TotOdMB    = Get-Stat $totRam 'avg'
            TetoWS     = Get-Mode $r 'SyncWsTetoMB'
            SyncOps    = Get-Stat (Col $r 'SyncOpsS') 'avg'
            OdOps      = Get-Stat (Col $r 'OdOpsS') 'avg'
            AvOps      = Get-Stat (Col $r 'AntivirusOpsS', 'DefenderOpsS') 'avg'
            IdxOps     = Get-Stat (Col $r 'IndexOpsS') 'avg'
            SyncIO     = Get-Stat (Col $r 'SyncIoMBs') 'avg'
            OdIO       = Get-Stat (Col $r 'OdIoMBs') 'avg'
            AvIO       = Get-Stat (Col $r 'AntivirusIoMBs', 'DefenderIoMBs') 'avg'
            IdxIO      = Get-Stat (Col $r 'IndexIoMBs') 'avg'
        }
    }
    $summary = @($summary | Sort-Object Dia, Config, Rotulo, @{ e = { Get-PhaseOrder $_.Fase } })

    # Quem fazia I/O nos minutos LENTOS (latencia media >= SlowLatMs).
    $lent = foreach ($g in ($rows | Where-Object $slow | Group-Object { $_.Hora.Substring(0, 10) }, Config, Fase)) {
        $r = $g.Group
        [pscustomobject]@{
            Dia       = $r[0].Hora.Substring(0, 10)
            Config    = $r[0].Config
            Fase      = $r[0].Fase
            MinLentos = $r.Count
            LatMs     = Get-Stat (Col $r 'DiscoLatMs') 'avg'
            SyncOps   = Get-Stat (Col $r 'SyncOpsS') 'avg'
            OdOps     = Get-Stat (Col $r 'OdOpsS') 'avg'
            AvOps     = Get-Stat (Col $r 'AntivirusOpsS', 'DefenderOpsS') 'avg'
            IdxOps    = Get-Stat (Col $r 'IndexOpsS') 'avg'
            SyncIO    = Get-Stat (Col $r 'SyncIoMBs') 'avg'
            AvIO      = Get-Stat (Col $r 'AntivirusIoMBs', 'DefenderIoMBs') 'avg'
            IdxIO     = Get-Stat (Col $r 'IndexIoMBs') 'avg'
            PgRd      = Get-Stat (Col $r 'PageReadsS') 'avg'
            TopFora   = Get-Mode $r 'TopIoProc'
            TopOps    = Get-Mode $r 'TopOpsProc'
        }
    }
    $lent = @($lent | Sort-Object Dia, Config, @{ e = { Get-PhaseOrder $_.Fase } })

    # Travamento x carga: minutos COM e SEM janela travada, lado a lado. Se nos minutos com trava
    # a latencia/CPU/operacoes do OneDrive sobem, o travamento acompanha a carga; se nao, a causa
    # e outra (ex.: o programa esperando o OneDrive/rede, sem carga de disco).
    $trv = foreach ($g in ($rows | Where-Object $measured |
                           Group-Object { $_.Hora.Substring(0, 10) }, Config, Fase, { if ((N $_.TravaAmostras) -gt 0) { 'com trava' } else { 'sem trava' } })) {
        $r = $g.Group
        [pscustomobject]@{
            Dia      = $r[0].Hora.Substring(0, 10)
            Config   = $r[0].Config
            Fase     = $r[0].Fase
            Minutos  = $g.Values[3]
            Min      = $r.Count
            LatMs    = Get-Stat (Col $r 'DiscoLatMs') 'avg'
            LatP95   = Get-Stat (Col $r 'DiscoLatMs') 'p95'
            CpuMed   = Get-Stat (Col $r 'CpuTotalPct') 'avg'
            CpuPico  = Get-Stat (Col $r 'CpuTotalMax') 'avg'
            SyncOps  = Get-Stat (Col $r 'SyncOpsS') 'avg'
            OdOps    = Get-Stat (Col $r 'OdOpsS') 'avg'
            AvOps    = Get-Stat (Col $r 'AntivirusOpsS') 'avg'
            IdxOps   = Get-Stat (Col $r 'IndexOpsS') 'avg'
            'PagUso%' = Get-Stat (Col $r 'PaginacaoUsoPct') 'avg'
            LivreMB  = Get-Stat (Col $r 'RamLivreMB') 'avg'
        }
    }
    $trv = @($trv | Sort-Object Dia, Config, @{ e = { Get-PhaseOrder $_.Fase } }, Minutos)

    # Quem travou: por dia e config, em quantos minutos e amostras cada processo ficou "Nao respondendo".
    $who = foreach ($g in ($rows | Where-Object { $_.TravaProcs } | Group-Object { $_.Hora.Substring(0, 10) }, Config)) {
        $g.Group | ForEach-Object { $_.TravaProcs -split ' / ' } | ForEach-Object {
            $i = $_.LastIndexOf(' ')
            [pscustomobject]@{ Proc = $_.Substring(0, $i); N = [int]$_.Substring($i + 1) }
        } | Group-Object Proc | ForEach-Object {
            [pscustomobject]@{
                Dia      = $g.Group[0].Hora.Substring(0, 10)
                Config   = $g.Group[0].Config
                Processo = $_.Name
                Minutos  = $_.Count
                Amostras = ($_.Group | Measure-Object N -Sum).Sum
                'Seg ~'  = ($_.Group | Measure-Object N -Sum).Sum * $DiskInterval
            }
        } | Sort-Object Minutos -Descending | Select-Object -First 10
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

    Write-Host ("`n== Resumo por dia, configuracao e fase (horas {0}-{1}) ==" -f $FromHour, $ToHour) -ForegroundColor Cyan
    Write-Host ("Fase: inicio = primeiros {0} min apos cada inicio do OneDrive | resto = depois disso | sem sync = motor de sync parado" -f $StartMinutes)
    Write-Host ('PROBLEMA REAL: MinTrava = minutos com alguma janela "Nao respondendo"; Trava% = % das amostras de {0} s com janela travada' -f $DiskInterval)
    Write-Host ("Disco: LatMs / LatP95 = latencia (media / p95, ms). MinLentos = minutos com latencia >= {0} ms. CpuMed/CpuP95 = CPU total da maquina" -f $SlowLatMs)
    Write-Host 'Ocup% e Em100% (so referencia em SSD) ficam no CSV do resumo. MinTrava vazio = CSV de versao sem medicao de travamento.'
    if ($estimated) {
        Write-Host ("(!) {0} linha(s) de CSV antigo sem SyncInicio: inicio do OneDrive estimado pela 1a aparicao do PID" -f $estimated) -ForegroundColor Yellow
    }
    # Tabelas separadas para caber no console; o CSV do resumo leva todas as colunas.
    Write-Host ($summary | Format-Table Dia, Config, Rotulo, Fase, Ativo, Min, MinTrava, 'Trava%', LatMs, LatP95, MinLentos, CpuMed, CpuP95, Iops -AutoSize |
        Out-String -Width 400).TrimEnd()
    Write-Host "`nMemoria: PagUso% / PagUsoMB = uso do arquivo de paginacao (media / maximo) - com LivreMB baixo, e falta de RAM."
    Write-Host 'PgRd = Page Reads/s: leituras de disco por falta de pagina, INCLUI arquivos fora do cache (nao e so paginacao).'
    Write-Host 'SyncPrivMB = memoria privada do sync (custo real); SyncMB/OdMB/TotOdMB = working set (o Windows reduz sozinho).'
    Write-Host ($summary | Format-Table Dia, Config, Fase, 'PagUso%', PagUsoMB, LivreMB, PgRd, SyncPrivMB, SyncMB, OdMB, TotOdMB, TetoWS -AutoSize |
        Out-String -Width 400).TrimEnd()
    Write-Host "`nI/O por grupo (Ops = operacoes/s, incluindo abrir/listar/ler atributos; IO = MB/s). Av = antivirus (ESET + Defender):"
    Write-Host ($summary | Format-Table Dia, Config, Fase, SyncOps, OdOps, AvOps, IdxOps, SyncIO, OdIO, AvIO, IdxIO -AutoSize |
        Out-String -Width 400).TrimEnd()

    Write-Host "`n== Travamentos: quem ficou 'Nao respondendo' ==" -ForegroundColor Cyan
    if ($who) { Write-Host ($who | Format-Table * -AutoSize | Out-String -Width 400).TrimEnd() }
    else      { Write-Host '  (nenhuma janela travada nas linhas medidas)' }
    Write-Host ("Minutos = minutos com o processo travado; Amostras x {0} s = tempo aproximado travado" -f $DiskInterval)

    Write-Host "`n== Travamento x carga: minutos com e sem janela travada ==" -ForegroundColor Cyan
    if ($trv) { Write-Host ($trv | Format-Table * -AutoSize | Out-String -Width 400).TrimEnd() }
    else      { Write-Host '  (sem dados - CSV de versao sem medicao de travamento)' }
    Write-Host 'Se latencia/CPU/Ops so sobem nos minutos "com trava", o travamento acompanha a carga. Se ficam iguais,'
    Write-Host 'o programa trava esperando outra coisa (ex.: o OneDrive ou a rede responderem), nao por disco/CPU cheios.'

    Write-Host ("`n== Nos minutos lentos (latencia >= {0} ms): quem fazia I/O ==" -f $SlowLatMs) -ForegroundColor Cyan
    if ($lent) { Write-Host ($lent | Format-Table * -AutoSize | Out-String -Width 400).TrimEnd() }
    else       { Write-Host '  (nenhum minuto lento)' }
    Write-Host 'TopFora / TopOps = processo de fora dos grupos que mais fez I/O (MB/s / operacoes) nesses minutos'

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
    $prevCpu  = Get-CpuRaw
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

        # --- CPU total e janelas "Nao respondendo", na mesma cadencia ---
        $curCpu = Get-CpuRaw
        $cpu = Get-CpuBusy $prevCpu $curCpu
        $prevCpu = $curCpu
        $hung = @(Get-HungWindows)

        $win.Add([pscustomobject]@{
            Od    = if ($od) { $od.Busy } else { $null }
            Queue = if ($od) { $od.Queue } else { $null }
            Iops  = if ($od) { $od.Iops } else { $null }
            LatMs = if ($od) { $od.LatMs } else { $null }
            Any   = if ($busy) { ($busy | Measure-Object Busy -Maximum).Maximum } else { $null }
            Cpu   = $cpu
            Hung  = $hung
        })

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
            # varredura de pastas e metadados, que e boa parte do trabalho do OneDrive/antivirus/indexador.
            $dOps = [double]$x.IODataOperationsPersec + [double]$x.IOOtherOperationsPersec -
                    $(if ($o) { [double]$o.IODataOperationsPersec + [double]$o.IOOtherOperationsPersec } else { 0 })
            if (-not $agg.ContainsKey($name)) {
                $agg[$name] = [pscustomobject]@{ Io = 0.0; Ops = 0.0; Cpu = 0.0; Ws = 0.0; Priv = 0.0; Pf = 0.0 }
            }
            $a = $agg[$name]
            $a.Io   += [math]::Max(0.0, [double]$dIo) / $secs
            $a.Ops  += [math]::Max(0.0, [double]$dOps) / $secs
            $a.Cpu  += [math]::Max(0.0, [double]$dCpu) / ($secs * 1e7) / $ncpu * 100
            $a.Pf   += [math]::Max(0.0, [double]$dPf) / $secs
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
        $syncCpuPrio = $null; $syncIo = $null; $syncWs = $null; $syncStart = $null; $ativo = 'n/a'
        if ($sp) {
            try { $syncCpuPrio = [string]$sp.PriorityClass } catch {}
            # Inicio do processo: o -Report separa a primeira hora (sincronizacao inicial) do resto do dia.
            try { $syncStart = $sp.StartTime.ToString('yyyy-MM-dd HH:mm:ss') } catch {}
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
        $anyW = @($win | Where-Object { $null -ne $_.Any })
        $aP95 = if ($anyW) { 100 * @($anyW | Where-Object { $_.Any -ge 95 }).Count / $anyW.Count } else { $null }

        # CPU total: media e pico das amostras de DiskInterval s.
        $cpuW   = @($win | Where-Object { $null -ne $_.Cpu })
        $cpuAvg = if ($cpuW) { ($cpuW | Measure-Object Cpu -Average).Average } else { $null }
        $cpuMax = if ($cpuW) { ($cpuW | Measure-Object Cpu -Maximum).Maximum } else { $null }

        # Travamentos: em quantas amostras havia janela "Nao respondendo" e quais processos.
        # TravaProcs = "nome amostras / nome amostras", do que travou mais vezes no minuto.
        $hungW     = @($win | Where-Object { $_.Hung.Count -gt 0 })
        $hungMax   = if ($win.Count) { ($win | ForEach-Object { $_.Hung.Count } | Measure-Object -Maximum).Maximum } else { 0 }
        $hungProcs = (@($hungW | ForEach-Object { $_.Hung | Select-Object -Unique }) | Group-Object |
            Sort-Object Count -Descending | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }) -join ' / '
        $dIops = if ($odW) { ($odW | Measure-Object Iops -Average).Average } else { $null }
        # Latencia media ponderada pelas operacoes de cada janela.
        $sumN  = ($odW | Measure-Object Iops -Sum).Sum
        $dLat  = if ($sumN -gt 0) { (($odW | ForEach-Object { $_.Iops * $_.LatMs }) | Measure-Object -Sum).Sum / $sumN } else { $null }

        $pr = ([double]$curMem.PageReadsPersec - [double]$prevMem.PageReadsPersec) / $secs
        $pg = Get-PagingFile
        $S = $grp['Sync']; $O = $grp['Od']; $X = $grp['OdOutros']; $A = $grp['Antivirus']; $I = $grp['Index']

        $row = [pscustomobject][ordered]@{
            Hora            = $now.ToString('yyyy-MM-dd HH:mm:ss')
            Rotulo          = $Label
            Config          = Get-ConfigTag
            AjusteAtivo     = $ativo
            SyncPid         = if ($sp) { $sp.Id } else { $null }
            SyncInicio      = $syncStart
            # Janelas "Nao respondendo" (o problema que o usuario sente).
            Amostras        = $win.Count
            TravaAmostras   = $hungW.Count
            TravaJanelasMax = [int]$hungMax
            TravaProcs      = $hungProcs
            CpuTotalPct     = if ($null -ne $cpuAvg) { Rnd $cpuAvg } else { $null }
            CpuTotalMax     = if ($null -ne $cpuMax) { Rnd $cpuMax } else { $null }
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
            # Leituras de disco por falta de pagina: arquivo de paginacao + arquivos fora do cache/mapeados.
            PageReadsS      = Rnd $pr
            PaginacaoUsoPct = if ($pg) { Rnd $pg.Pct } else { $null }
            PaginacaoUsoMB  = if ($pg) { [int]$pg.MB } else { $null }
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
            AntivirusIoMBs  = Rnd ($A.Io / 1MB) 2
            AntivirusOpsS   = Rnd $A.Ops 0
            AntivirusCpuPct = Rnd $A.Cpu
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

        $travaTxt = if ($row.TravaAmostras) { 'TRAVA {0}/{1}: {2}' -f $row.TravaAmostras, $row.Amostras, $row.TravaProcs } else { 'sem trava' }
        Write-Host ("{0}  {1}  |  latencia {2,6} ms  {3,5} IOPS  cpu {4,5}% (pico {5})  sync priv {6,5} MB  pagefile {7}%  livre {8} MB  [{9}]" -f
            $now.ToString('HH:mm'), $travaTxt, $row.DiscoLatMs, $row.DiscoIops, $row.CpuTotalPct, $row.CpuTotalMax, $row.SyncPrivMB,
            $row.PaginacaoUsoPct, $row.RamLivreMB, $row.Config) -ForegroundColor $(if ($row.TravaAmostras) { 'Red' } else { 'Gray' })
    }
} finally {
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
