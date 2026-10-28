# Ставит бота службой Windows через NSSM (как в соседнем проекте cashflow1c).
# Запускать НА СЕРВЕРЕ в PowerShell ОТ ИМЕНИ АДМИНИСТРАТОРА:
#
#     powershell -ExecutionPolicy Bypass -File C:\daily-report\tools\install_service_nssm.ps1 -Pass "пароль"
#
# ВАЖНО (проверено в cashflow1c): служба работает в сеансе 0, где нет
# рабочего стола. Автозапуск самой 1С оттуда не сработает — на сервере
# тонкий клиент должен быть уже открыт в обычном RDP-сеансе под той же
# учётной записью ($User), тогда бот подключается к его окну.

param(
    [string]$Project = "C:\daily-report",
    [string]$User = ".\Administrator",
    [Parameter(Mandatory = $true)][string]$Pass,
    [string]$Service = "dailysales-bot",
    [string]$Nssm = "C:\nssm\nssm.exe",
    [string]$Python
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path "$Project\bot.py")) { throw "Не нашёл $Project\bot.py" }

if (-not $Python) { $Python = (Get-Command python).Source }
if (-not (Test-Path $Python)) { throw "Не нашёл python.exe: $Python" }

if (-not (Test-Path $Nssm)) {
    Write-Host "Скачиваю NSSM..." -ForegroundColor Cyan
    $dir = Split-Path -Parent $Nssm
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest "https://nssm.cc/release/nssm-2.24.zip" -OutFile "$dir\nssm.zip"
    Expand-Archive "$dir\nssm.zip" -DestinationPath $dir -Force
    Copy-Item "$dir\nssm-2.24\win64\nssm.exe" $Nssm -Force
}

& $Nssm status $Service 2>$null | Out-Null
if ($LASTEXITCODE -eq 0) {
    Write-Host "Служба уже есть — переустанавливаю..." -ForegroundColor Yellow
    & $Nssm stop $Service | Out-Null
    & $Nssm remove $Service confirm | Out-Null
    Start-Sleep -Seconds 2
}

New-Item -ItemType Directory -Force -Path "$Project\logs" | Out-Null

& $Nssm install $Service $Python "$Project\bot.py"
& $Nssm set $Service AppDirectory $Project
& $Nssm set $Service DisplayName "Daily Sales bot (1C -> Google Sheets)"
& $Nssm set $Service Description "Ежедневный отчёт: 1С -> лист Регистр -> скриншот Sales by day -> Telegram"
& $Nssm set $Service AppStdout "$Project\logs\bot.log"
& $Nssm set $Service AppStderr "$Project\logs\bot-error.log"
& $Nssm set $Service AppRotateFiles 1
& $Nssm set $Service AppRotateBytes 10485760
& $Nssm set $Service AppEnvironmentExtra PYTHONIOENCODING=utf-8 VIRTUAL_INPUT_ALLOW_REAL=0
& $Nssm set $Service Start SERVICE_AUTO_START
& $Nssm set $Service AppExit Default Restart
& $Nssm set $Service AppRestartDelay 10000
& $Nssm set $Service ObjectName $User $Pass

Start-Service $Service
Start-Sleep -Seconds 6
Get-Service $Service | Format-Table Name, Status -AutoSize
Get-Content "$Project\logs\bot.log" -Tail 15 -ErrorAction SilentlyContinue

Write-Host "`nУправление:" -ForegroundColor Cyan
Write-Host "  Restart-Service $Service"
Write-Host "  Get-Content $Project\logs\bot.log -Tail 30 -Wait"
Write-Host "  & $Nssm remove $Service confirm   # снять службу"
