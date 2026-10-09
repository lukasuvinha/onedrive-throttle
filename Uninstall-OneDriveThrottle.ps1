<#
.SYNOPSIS
    Remove tudo o que o onedrive-throttle instala nesta maquina, em um passo.

.DESCRIPTION
    Desfaz, nesta ordem:
      1. Ajuste de prioridade e politicas (Set-OneDriveThrottle.ps1 -Action Remove):
         chaves IFEO PerfOptions dos processos do OneDrive e as politicas que o
         projeto pode ter criado.
      2. Tarefa agendada "onedrive-throttle-watch" e o lancador Watch-OneDrive.vbs.
      3. Atalho antigo Watch-OneDrive.lnk da pasta Inicializar de TODOS os perfis.
      4. Monitores em execucao (powershell.exe rodando Watch-OneDrive.ps1).
      5. Opcional (-DeleteReports): a pasta reports\ ao lado dos scripts.
      6. Opcional (-CollectDir): a subpasta <PC>_<USUARIO> desta maquina na pasta
         de coleta central.

    Nao apaga a pasta do projeto (os proprios scripts). Depois de rodar, apague a
    pasta manualmente se quiser remover tudo.

    Precisa de PowerShell como administrador. Depois: logoff/login ou reiniciar,
    para o OneDrive voltar a iniciar com prioridade normal.

.PARAMETER DeleteReports
    Apaga tambem a pasta reports\ (CSVs, resumos, diagnosticos e log).

.PARAMETER CollectDir
    Mesma pasta de coleta usada no -InstallTask (absoluta ou relativa ao OneDrive
    do usuario). Se informada, apaga a subpasta desta maquina e deste usuario.

.PARAMETER KeepThrottle
    Remove so o monitor e mantem o ajuste de prioridade do OneDrive.

.EXAMPLE
    .\Uninstall-OneDriveThrottle.ps1
    Remove ajuste, tarefa, atalhos e para o monitor. Mantem os relatorios.

.EXAMPLE
    .\Uninstall-OneDriveThrottle.ps1 -DeleteReports -CollectDir 'Pasta\Subpasta'
    Remove tudo, inclusive os relatorios locais e a copia desta maquina na coleta.
#>

[CmdletBinding()]
param(
    [switch]$DeleteReports,
    [string]$CollectDir = '',
    [switch]$KeepThrottle
)

$ErrorActionPreference = 'Continue'
$here = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
$TaskName = 'onedrive-throttle-watch'

function Step([string]$t) { Write-Host "`n== $t ==" -ForegroundColor Cyan }
function Ok([string]$t)   { Write-Host "  [ok] $t" -ForegroundColor Green }
function Info([string]$t) { Write-Host "  $t" }
function Warn([string]$t) { Write-Host "  [!] $t" -ForegroundColor Yellow }

$id = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not (New-Object Security.Principal.WindowsPrincipal $id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host 'Execute como administrador (botao direito no PowerShell > Executar como administrador).' -ForegroundColor Red
    return
}

# 1. Ajuste de prioridade e politicas
Step '1. Ajuste de prioridade do OneDrive'
if ($KeepThrottle) {
    Info 'Mantido (-KeepThrottle).'
} else {
    $set = Join-Path $here 'Set-OneDriveThrottle.ps1'
    if (Test-Path -LiteralPath $set) {
        & $set -Action Remove
        Ok 'Set-OneDriveThrottle.ps1 -Action Remove executado.'
    } else {
        # Sem o script: remove direto as chaves IFEO conhecidas.
        Warn 'Set-OneDriveThrottle.ps1 nao encontrado; removendo as chaves IFEO diretamente.'
        $exes = 'OneDrive.exe', 'OneDrive.Sync.Service.exe', 'FileCoAuth.exe', 'Microsoft.SharePoint.exe', 'OneDriveStandaloneUpdater.exe'
        $roots = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options',
                 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows NT\CurrentVersion\Image File Execution Options'
        foreach ($r in $roots) { foreach ($e in $exes) {
            $po = Join-Path $r "$e\PerfOptions"
            if (Test-Path $po) { Remove-Item $po -Recurse -Force; Ok "Removido $po" }
            $k = Join-Path $r $e
            if ((Test-Path $k) -and -not (Get-ChildItem $k) -and -not (Get-Item $k).ValueCount) { Remove-Item $k -Force }
        } }
    }
}

# 2. Tarefa agendada e lancador
Step '2. Tarefa agendada do monitor'
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Ok "Tarefa '$TaskName' removida."
} else { Info "Tarefa '$TaskName' nao existia." }
$vbs = Join-Path $here 'Watch-OneDrive.vbs'
if (Test-Path -LiteralPath $vbs) { Remove-Item -LiteralPath $vbs -Force; Ok "Lancador removido: $vbs" }

# 3. Atalhos antigos na pasta Inicializar de todos os perfis
Step '3. Atalhos na pasta Inicializar'
$found = 0
$profiles = Get-CimInstance Win32_UserProfile -ErrorAction SilentlyContinue | Where-Object { -not $_.Special -and $_.LocalPath }
foreach ($p in $profiles) {
    $lnk = Join-Path $p.LocalPath 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Watch-OneDrive.lnk'
    if (Test-Path -LiteralPath $lnk) { Remove-Item -LiteralPath $lnk -Force; Ok "Removido $lnk"; $found++ }
}
if (-not $found) { Info 'Nenhum atalho encontrado.' }

# 4. Monitores em execucao
Step '4. Monitores em execucao'
$procs = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -like '*Watch-OneDrive.ps1*' -and $_.ProcessId -ne $PID })
if ($procs) {
    foreach ($p in $procs) {
        try { Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop; Ok "Monitor encerrado (PID $($p.ProcessId))." }
        catch { Warn "Nao consegui encerrar o PID $($p.ProcessId): $($_.Exception.Message)" }
    }
} else { Info 'Nenhum monitor rodando.' }

# 5. Relatorios locais
Step '5. Relatorios locais'
$reports = Join-Path $here 'reports'
if ($DeleteReports) {
    if (Test-Path -LiteralPath $reports) { Remove-Item -LiteralPath $reports -Recurse -Force; Ok "Apagado $reports" }
    else { Info 'Pasta reports nao existia.' }
} else { Info "Mantidos em $reports (use -DeleteReports para apagar)." }

# 6. Copia desta maquina na coleta central
if ($CollectDir) {
    Step '6. Coleta central'
    $base = $CollectDir
    if (-not [IO.Path]::IsPathRooted($base)) {
        $od = if ($env:OneDriveCommercial) { $env:OneDriveCommercial } else { $env:OneDrive }
        if ($od) { $base = Join-Path $od $base } else { $base = $null }
    }
    if ($base) {
        $mine = Join-Path $base ("{0}_{1}" -f $env:COMPUTERNAME, $env:USERNAME)
        if (Test-Path -LiteralPath $mine) { Remove-Item -LiteralPath $mine -Recurse -Force; Ok "Apagado $mine" }
        else { Info "Nao encontrei $mine" }
    } else { Warn 'OneDrive do usuario nao encontrado; apague a subpasta manualmente.' }
}

Write-Host "`nPronto. Faca logoff/login ou reinicie para o OneDrive voltar a prioridade normal." -ForegroundColor Green
Write-Host "Para apagar os scripts, exclua a pasta: $here"
