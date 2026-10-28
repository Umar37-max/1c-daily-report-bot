# Установка на сервере. Запускать НА СЕРВЕРЕ, в том сеансе (RDP), под
# которым будет работать 1С:
#
#     powershell -ExecutionPolicy Bypass -File tools\install_on_server.ps1
#
# Делает: проверяет Python и 1С, ставит зависимости, прогоняет проверку
# настроек, ставит автозапуск бота при входе в систему и запускает его.

param(
    [string]$Python = "python",
    [switch]$SkipDeps
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

Write-Output "=== 1/5 Python ==="
$version = & $Python --version 2>&1
if ($LASTEXITCODE -ne 0) { throw "Python не найден. Поставь Python 3.11+ и добавь в PATH." }
Write-Output $version

Write-Output "`n=== 2/5 Зависимости ==="
if ($SkipDeps) {
    Write-Output "пропущено (-SkipDeps)"
} else {
    & $Python -m pip install --disable-pip-version-check -q -r requirements.txt
    if ($LASTEXITCODE -ne 0) { throw "pip install не прошёл" }
    Write-Output "установлены"
}

Write-Output "`n=== 3/5 Тонкий клиент 1С ==="
$clients = @()
foreach ($base in @("C:\Program Files\1cv8", "C:\Program Files (x86)\1cv8")) {
    if (Test-Path $base) {
        $clients += Get-ChildItem $base -Directory |
            Where-Object { Test-Path (Join-Path $_.FullName "bin\1cv8c.exe") } |
            ForEach-Object { $_.Name }
    }
}
if ($clients.Count -eq 0) {
    Write-Warning "1cv8c.exe не найден — на сервере должен стоять тонкий клиент 1С."
} else {
    Write-Output ("найдены версии: " + ($clients -join ", "))
}
$ibases = "$env:APPDATA\1C\1CEStart\ibases.v8i"
if (Test-Path $ibases) {
    Write-Output "список баз этой учётки: $ibases"
} else {
    Write-Warning "У этой учётки нет списка баз 1С — подключение пойдёт по ONEC_CONNECTION_STRING из .env (так и задумано)."
}

Write-Output "`n=== 4/5 Проверка настроек ==="
$env:PYTHONIOENCODING = "utf-8"
& $Python tools\check_setup.py
if ($LASTEXITCODE -ne 0) {
    Write-Warning "Проверка нашла замечания (см. выше). Автозапуск всё равно ставлю — поправь .env и перезапусти бота."
}

Write-Output "`n=== 5/5 Автозапуск бота ==="
& powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "setup_schedule.ps1")
Start-ScheduledTask -TaskName "DailySalesBot"
Start-Sleep -Seconds 5
$running = Get-Process -Name python -ErrorAction SilentlyContinue
if ($running) {
    Write-Output "Бот запущен."
} else {
    Write-Warning "Процесс python не виден — смотри logs\bot.log"
}

Write-Output @"

Готово. Дальше:
  * сеанс должен оставаться залогиненным — закрывай RDP кнопкой «Отключиться»,
    а не «Выйти из системы», иначе рабочего стола не будет и 1С не откроется;
  * логи бота: logs\bot.log, логи прогонов: logs\run_*.log;
  * управление — в Telegram: /run, /week, /date, /shot, /status;
  * автозапуск каждый день в RUN_AT (.env).
"@
