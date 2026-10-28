# Заливка проекта на сервер.
#
# Вариант 1 — просто собрать архив и отнести руками:
#     powershell -ExecutionPolicy Bypass -File tools\deploy_to_server.ps1 -Zip
#
# Вариант 2 — скопировать по сети на уже смонтированный/доступный путь:
#     powershell -ExecutionPolicy Bypass -File tools\deploy_to_server.ps1 -Target \\SERVER\c$\daily-report
#
# Вариант 3 — по WinRM в сессию сервера (нужны права и включённый WinRM):
#     powershell -ExecutionPolicy Bypass -File tools\deploy_to_server.ps1 -ComputerName SRV1 -Target C:\daily-report
#
# Что копируется: код, .env и credentials\ (без них на сервере ничего не
# заработает). Не копируются logs\, out\, __pycache__, tools\_1c_export.
#
# После копирования на СЕРВЕРЕ выполнить:
#     powershell -ExecutionPolicy Bypass -File tools\install_on_server.ps1

param(
    [string]$Target,
    [string]$ComputerName,
    [switch]$Zip
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$exclude = @("logs", "out", "__pycache__", ".git", "_1c_export")

function New-Package {
    $stage = Join-Path $env:TEMP "daily-report-deploy"
    if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
    New-Item -ItemType Directory -Path $stage | Out-Null

    robocopy $root $stage /E /XD $exclude /XF "*.pyc" "bot_test.*" /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy завершился с кодом $LASTEXITCODE" }
    return $stage
}

$stage = New-Package
Write-Output "Собрал комплект: $stage"

if ($Zip) {
    $zipPath = Join-Path $root ("daily-report_{0:yyyyMMdd_HHmm}.zip" -f (Get-Date))
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $zipPath
    Write-Output "Архив: $zipPath"
    Write-Output "Распакуй его на сервере (например в C:\daily-report) и запусти tools\install_on_server.ps1"
    exit 0
}

if (-not $Target) { throw "Укажи -Target (куда копировать) или -Zip" }

if ($ComputerName) {
    Write-Output "Подключаюсь к $ComputerName..."
    $session = New-PSSession -ComputerName $ComputerName
    try {
        Invoke-Command -Session $session -ScriptBlock {
            param($path)
            if (-not (Test-Path $path)) { New-Item -ItemType Directory -Path $path -Force | Out-Null }
        } -ArgumentList $Target
        Copy-Item -Path (Join-Path $stage "*") -Destination $Target -ToSession $session -Recurse -Force
        Write-Output "Скопировано в $ComputerName : $Target"
        Write-Output "Теперь на сервере: powershell -ExecutionPolicy Bypass -File $Target\tools\install_on_server.ps1"
    }
    finally {
        Remove-PSSession $session
    }
}
else {
    if (-not (Test-Path $Target)) { New-Item -ItemType Directory -Path $Target -Force | Out-Null }
    robocopy $stage $Target /E /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy завершился с кодом $LASTEXITCODE" }
    Write-Output "Скопировано в $Target"
    Write-Output "Теперь на сервере: powershell -ExecutionPolicy Bypass -File $Target\tools\install_on_server.ps1"
}
