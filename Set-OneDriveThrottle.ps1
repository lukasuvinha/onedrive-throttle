<#
.SYNOPSIS
    Reduz o impacto do OneDrive (CPU, disco, rede e espaco) no Windows 10/11
    usando apenas configuracoes nativas do registro. Nao fica rodando em
    segundo plano: aplica uma vez e o Windows cuida do resto a cada boot.

.DESCRIPTION
    Duas camadas:

    1) PRIORIDADE DE PROCESSO (IFEO PerfOptions)
       Grava em HKLM\...\Image File Execution Options\<exe>\PerfOptions a
       prioridade de CPU e de I/O. O Windows aplica isso no momento em que o
       processo nasce - toda vez, inclusive depois de reboot, update ou crash
       do OneDrive. Substitui o "script no startup que ajusta prioridade".

    2) POLITICAS OFICIAIS DO ONEDRIVE (HKLM\SOFTWARE\Policies\Microsoft\OneDrive)
       Files On-Demand, limite de upload, exclusao de tipos de arquivo, etc.
       Sao as mesmas chaves que o GPO/ADMX do OneDrive grava.

    Tudo e reversivel com -Action Remove.

.PARAMETER Action
    Status (padrao) - mostra o que esta configurado e a prioridade atual dos processos.
    Apply           - aplica as configuracoes.
    Remove          - desfaz tudo que este script cria.

.PARAMETER CpuPriority
    BelowNormal (padrao, recomendado) ou Idle.
    Idle pode fazer a sincronizacao "nunca terminar" em PC sempre ocupado.

.PARAMETER IoPriority
    Low (padrao) ou VeryLow. Reduz a disputa de disco - o maior ganho em HD mecanico.

.PARAMETER LowMemoryPriority
    EXPERIMENTAL. Define PagePriority baixa: as paginas do OneDrive sao as
    primeiras a sair da RAM quando falta memoria. NAO limita o consumo de RAM.

.PARAMETER MaxWorkingSetMB
    EXPERIMENTAL. Teto rigido de RAM fisica (working set) do OneDrive.exe, via
    IFEO WorkingSetLimitInKB. O que passar do teto vai para o arquivo de
    paginacao - ou seja, RAM economizada pode virar LEITURA DE DISCO.
    Teste com Get-OneDriveDiag.ps1 antes e depois (olhe "Page Reads/s").

.PARAMETER FilesOnDemand
    Forca o Files On-Demand (arquivos so baixam quando abertos). Economiza disco.

.PARAMETER DehydrateTeamSites
    Converte arquivos de bibliotecas do SharePoint/Teams sincronizadas para
    "somente online". Requer -FilesOnDemand.

.PARAMETER UploadPercent
    Limita o upload a uma % da banda (10-99). A Microsoft recomenda >= 50.

.PARAMETER AutoUploadBandwidth
    Usa o gerenciamento automatico de banda (LEDBAT): so sobe arquivo com banda ociosa.
    Nao combine com -UploadPercent (este tem prioridade na propria Microsoft).

.PARAMETER IgnorePatterns
    Padroes de arquivo que o OneDrive nao deve enviar. Ex.: '*.pst','*.ldb'
    So vale para arquivos NOVOS. Exige reiniciar o OneDrive.

.PARAMETER DehydrateAfterDays
    Storage Sense: transforma em "somente online" arquivos da nuvem nao abertos
    ha N dias (1, 14, 30 ou 60). Requer Files On-Demand.

.EXAMPLE
    .\Set-OneDriveThrottle.ps1
    Mostra o estado atual.

.EXAMPLE
    .\Set-OneDriveThrottle.ps1 -Action Apply
    Aplica so a prioridade (CPU BelowNormal + I/O Low).

.EXAMPLE
    .\Set-OneDriveThrottle.ps1 -Action Apply -FilesOnDemand -AutoUploadBandwidth -IgnorePatterns '*.pst','*.ldb','*.laccdb' -DehydrateAfterDays 30
    Aplica prioridade + politicas.

.EXAMPLE
    .\Set-OneDriveThrottle.ps1 -Action Remove
    Desfaz tudo.

.NOTES
    Precisa de administrador para Apply/Remove.
    Efeito da prioridade: a partir do proximo inicio do OneDrive (logoff/login ou reboot).
    IFEO PerfOptions nao e documentado oficialmente pela Microsoft, mas funciona
    desde o Windows Vista e e amplamente usado.
#>

[CmdletBinding()]
param(
    [ValidateSet('Status', 'Apply', 'Remove')]
    [string]$Action = 'Status',

    [ValidateSet('BelowNormal', 'Idle')]
    [string]$CpuPriority = 'BelowNormal',

    [ValidateSet('Low', 'VeryLow')]
    [string]$IoPriority = 'Low',

    [switch]$LowMemoryPriority,

    [ValidateRange(200, 4096)]
    [int]$MaxWorkingSetMB,

    [switch]$FilesOnDemand,
    [switch]$DehydrateTeamSites,

    [ValidateRange(10, 99)]
    [int]$UploadPercent,

    [switch]$AutoUploadBandwidth,

    [string[]]$IgnorePatterns,

    [ValidateSet(1, 14, 30, 60)]
    [int]$DehydrateAfterDays
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

# Executaveis da familia OneDrive que recebem prioridade reduzida.
$Targets = @(
    'OneDrive.exe',                    # cliente principal (icone, interface)
    'OneDrive.Sync.Service.exe',       # motor de sincronizacao (builds 2025+) - costuma ser o mais pesado
    'FileCoAuth.exe',                  # coautoria com Office
    'Microsoft.SharePoint.exe',        # sincronizacao de bibliotecas SharePoint/Teams
    'OneDriveStandaloneUpdater.exe'    # atualizador
)

# Valores de PerfOptions (REG_DWORD)
$CpuMap = @{ 'Idle' = 1; 'BelowNormal' = 5 }   # 2 = Normal
$IoMap  = @{ 'VeryLow' = 0; 'Low' = 1 }        # 2 = Normal (IFEO nao aceita acima de Normal)
$PageLow = 2                                   # 5 = Normal; 1-2 = baixa

# Vistas do registro: nativa e 32 bits (Wow6432Node). Gravar nas duas cobre
# tanto o OneDrive 64 bits quanto instalacoes antigas de 32 bits.
$IfeoRoots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows NT\CurrentVersion\Image File Execution Options'
)

$OdPolicy  = 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'
$OdIgnore  = "$OdPolicy\EnableODIgnoreListFromGPO"
$SsPolicy  = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\StorageSense'

# Valores de politica que este script pode criar (usado pelo Remove).
$OdValues = @(
    'FilesOnDemandEnabled',
    'DehydrateSyncedTeamSites',
    'AutomaticUploadBandwidthPercentage',
    'EnableAutomaticUploadBandwidthManagement'
)
$SsValues = @(
    'AllowStorageSenseGlobal',
    'ConfigStorageSenseCloudContentDehydrationThreshold'
)

# ---------------------------------------------------------------------------
# Utilitarios
# ---------------------------------------------------------------------------

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal $id).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Set-Dword([string]$Path, [string]$Name, [int]$Value) {
    if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    New-ItemProperty -Path $Path -Name $Name -PropertyType DWord -Value $Value -Force | Out-Null
    Write-Host ("  [+] {0}\{1} = {2}" -f $Path, $Name, $Value)
}

function Remove-Values([string]$Path, [string[]]$Names) {
    if (-not (Test-Path $Path)) { return }
    foreach ($n in $Names) {
        if ($null -ne (Get-ItemProperty -Path $Path -Name $n -ErrorAction SilentlyContinue)) {
            Remove-ItemProperty -Path $Path -Name $n -Force
            Write-Host ("  [-] {0}\{1}" -f $Path, $n)
        }
    }
}

function Remove-KeyIfEmpty([string]$Path) {
    if (-not (Test-Path $Path)) { return }
    $item = Get-Item $Path
    if ($item.SubKeyCount -eq 0 -and $item.ValueCount -eq 0) {
        Remove-Item $Path -Force
    }
}

# ---------------------------------------------------------------------------
# Acoes
# ---------------------------------------------------------------------------

function Invoke-Apply {
    Write-Host "`n== Prioridade de processo (IFEO PerfOptions) ==" -ForegroundColor Cyan
    foreach ($root in $IfeoRoots) {
        if (-not (Test-Path (Split-Path $root))) { continue }   # sem Wow6432Node em SO 32 bits
        foreach ($exe in $Targets) {
            $po = Join-Path $root "$exe\PerfOptions"
            Set-Dword $po 'CpuPriorityClass' $CpuMap[$CpuPriority]
            Set-Dword $po 'IoPriority'       $IoMap[$IoPriority]
            if ($LowMemoryPriority) { Set-Dword $po 'PagePriority' $PageLow }
            # Teto de RAM nos dois processos grandes (vale POR PROCESSO, nao somado).
            if ($MaxWorkingSetMB -and $exe -in 'OneDrive.exe', 'OneDrive.Sync.Service.exe') {
                Set-Dword $po 'WorkingSetLimitInKB' ($MaxWorkingSetMB * 1024)
            }
        }
    }
    if ($MaxWorkingSetMB) {
        Write-Host '  (!) Teto de RAM e experimental: acompanhe Page Reads/s e o disco depois de aplicar.' -ForegroundColor Yellow
    }

    $anyPolicy = $FilesOnDemand -or $DehydrateTeamSites -or $UploadPercent -or
                 $AutoUploadBandwidth -or $IgnorePatterns -or $DehydrateAfterDays
    if (-not $anyPolicy) { return }

    Write-Host "`n== Politicas do OneDrive ==" -ForegroundColor Cyan
    if ($FilesOnDemand)       { Set-Dword $OdPolicy 'FilesOnDemandEnabled' 1 }
    if ($DehydrateTeamSites)  {
        if (-not $FilesOnDemand) { Write-Warning 'DehydrateTeamSites so funciona com Files On-Demand ativo.' }
        Set-Dword $OdPolicy 'DehydrateSyncedTeamSites' 1
    }
    if ($UploadPercent -and $AutoUploadBandwidth) {
        Write-Warning 'UploadPercent e AutoUploadBandwidth juntos: a Microsoft recomenda usar apenas um. Aplicando so AutoUploadBandwidth.'
        $script:UploadPercent = 0
    }
    if ($UploadPercent)       { Set-Dword $OdPolicy 'AutomaticUploadBandwidthPercentage' $UploadPercent }
    if ($AutoUploadBandwidth) { Set-Dword $OdPolicy 'EnableAutomaticUploadBandwidthManagement' 1 }

    if ($IgnorePatterns) {
        if (-not (Test-Path $OdIgnore)) { New-Item -Path $OdIgnore -Force | Out-Null }
        foreach ($p in $IgnorePatterns) {
            New-ItemProperty -Path $OdIgnore -Name $p -PropertyType String -Value $p -Force | Out-Null
            Write-Host ("  [+] {0}\{1}" -f $OdIgnore, $p)
        }
        Write-Host '  (!) A lista de exclusao so vale apos reiniciar o OneDrive e so para arquivos novos.' -ForegroundColor Yellow
    }

    if ($DehydrateAfterDays) {
        Write-Host "`n== Storage Sense ==" -ForegroundColor Cyan
        Set-Dword $SsPolicy 'AllowStorageSenseGlobal' 1
        Set-Dword $SsPolicy 'ConfigStorageSenseCloudContentDehydrationThreshold' $DehydrateAfterDays
    }
}

function Invoke-Remove {
    Write-Host "`n== Removendo prioridade (IFEO) ==" -ForegroundColor Cyan
    foreach ($root in $IfeoRoots) {
        foreach ($exe in $Targets) {
            $exeKey = Join-Path $root $exe
            $po     = Join-Path $exeKey 'PerfOptions'
            Remove-Values $po @('CpuPriorityClass', 'IoPriority', 'PagePriority', 'WorkingSetLimitInKB')
            Remove-KeyIfEmpty $po
            Remove-KeyIfEmpty $exeKey     # so apaga a chave do exe se ficou vazia
        }
    }

    Write-Host "`n== Removendo politicas ==" -ForegroundColor Cyan
    Remove-Values $OdPolicy $OdValues
    if (Test-Path $OdIgnore) {
        Remove-Item $OdIgnore -Recurse -Force
        Write-Host "  [-] $OdIgnore"
    }
    Remove-KeyIfEmpty $OdPolicy
    Remove-Values $SsPolicy $SsValues
    Remove-KeyIfEmpty $SsPolicy
}

function Show-Status {
    Write-Host "`n== IFEO PerfOptions ==" -ForegroundColor Cyan
    foreach ($root in $IfeoRoots) {
        foreach ($exe in $Targets) {
            $po = Join-Path $root "$exe\PerfOptions"
            if (Test-Path $po) {
                $v = Get-ItemProperty $po
                Write-Host ("  {0,-32} CPU={1} IO={2} Page={3} WSLimitKB={4}  ({5})" -f $exe,
                    $v.CpuPriorityClass, $v.IoPriority, $v.PagePriority, $v.WorkingSetLimitInKB,
                    $(if ($root -like '*WOW6432Node*') { '32 bits' } else { 'nativo' }))
            }
        }
    }

    Write-Host "`n== Politicas do OneDrive (HKLM) ==" -ForegroundColor Cyan
    if (Test-Path $OdPolicy) {
        (Get-ItemProperty $OdPolicy).PSObject.Properties |
            Where-Object { $_.Name -notlike 'PS*' } |
            ForEach-Object { Write-Host ("  {0} = {1}" -f $_.Name, $_.Value) }
    } else { Write-Host '  (nenhuma)' }
    if (Test-Path $OdIgnore) {
        $list = (Get-ItemProperty $OdIgnore).PSObject.Properties |
            Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object { $_.Value }
        Write-Host ("  Exclusoes: {0}" -f ($list -join ', '))
    }

    Write-Host "`n== Storage Sense (politica) ==" -ForegroundColor Cyan
    if (Test-Path $SsPolicy) {
        (Get-ItemProperty $SsPolicy).PSObject.Properties |
            Where-Object { $_.Name -notlike 'PS*' } |
            ForEach-Object { Write-Host ("  {0} = {1}" -f $_.Name, $_.Value) }
    } else { Write-Host '  (nenhuma)' }

    Write-Host "`n== Processos em execucao ==" -ForegroundColor Cyan
    $names = $Targets | ForEach-Object { [IO.Path]::GetFileNameWithoutExtension($_) }
    $procs = Get-Process -Name $names -ErrorAction SilentlyContinue
    if ($procs) {
        $procs | Select-Object Name, Id, PriorityClass,
            @{ n = 'RAM(MB)'; e = { [math]::Round($_.WorkingSet64 / 1MB) } },
            @{ n = 'CPU(s)';  e = { [math]::Round($_.CPU) } } |
            Format-Table -AutoSize | Out-String | Write-Host
    } else { Write-Host '  (nenhum processo do OneDrive rodando)' }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

switch ($Action) {
    'Status' { Show-Status }
    'Apply' {
        if (-not (Test-Admin)) { throw 'Execute como administrador.' }
        Invoke-Apply
        Write-Host "`nPronto. Vale a partir do proximo inicio do OneDrive (logoff/login ou reboot)." -ForegroundColor Green
    }
    'Remove' {
        if (-not (Test-Admin)) { throw 'Execute como administrador.' }
        Invoke-Remove
        Write-Host "`nRemovido. Reinicie o OneDrive (ou o PC) para voltar ao normal." -ForegroundColor Green
    }
}
