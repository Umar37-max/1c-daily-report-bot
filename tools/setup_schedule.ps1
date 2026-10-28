# Ставит автозапуск в Планировщик задач Windows.
#
#   # рекомендуемый режим: бот стартует при входе в систему и сам держит
#   # расписание 06:30 (и им же можно управлять из Telegram)
#   powershell -ExecutionPolicy Bypass -File tools\setup_schedule.ps1
#
#   # альтернатива без бота: просто запускать прогон каждый день в 06:30
#   powershell -ExecutionPolicy Bypass -File tools\setup_schedule.ps1 -Mode report -Time 06:30
#
#   powershell -ExecutionPolicy Bypass -File tools\setup_schedule.ps1 -Remove
#
# ВАЖНО: задачи ставятся как «выполнять только при входе пользователя».
# Служба Windows живёт в сеансе 0, где нет рабочего стола, — тонкий клиент
# 1С там не открывается и прогон падает всегда (проверено на сервере
# соседнего проекта). Сеанс должен оставаться залогиненным: RDP закрывать
# кнопкой «Отключиться», а не «Выйти из системы».
#
# Одновременно оба режима включать НЕ надо — прогон запустится дважды.

param(
    [ValidateSet("bot", "report")]
    [string]$Mode = "bot",
    [string]$Time = "06:30",
    [switch]$Remove
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$botTask = "DailySalesBot"
$reportTask = "DailySalesReport"

function Remove-IfExists($name) {
    if (Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $name -Confirm:$false
        Write-Output "Удалена задача '$name'."
    }
}

if ($Remove) {
    Remove-IfExists $botTask
    Remove-IfExists $reportTask
    exit 0
}

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
    -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

if ($Mode -eq "bot") {
    Remove-IfExists $reportTask
    $vbs = Join-Path $root "start_bot_hidden.vbs"
    if (-not (Test-Path $vbs)) { throw "Не найден $vbs" }

    $action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$vbs`"" -WorkingDirectory $root
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
    Register-ScheduledTask -TaskName $botTask -Action $action -Trigger $trigger `
        -Settings $settings -Principal $principal -Force `
        -Description "Телеграм-бот ежедневного отчёта по продажам (сам держит расписание $Time)" | Out-Null

    Write-Output "Задача '$botTask' создана: бот стартует при входе в систему."
    Write-Output "Расписание прогона (${Time}) держит сам бот — см. RUN_AT в .env."
    Write-Output "Запустить сейчас: Start-ScheduledTask -TaskName $botTask"
}
else {
    Remove-IfExists $botTask
    $runner = Join-Path $root "run_daily.cmd"
    if (-not (Test-Path $runner)) { throw "Не найден $runner" }

    $action = New-ScheduledTaskAction -Execute $runner -WorkingDirectory $root
    $trigger = New-ScheduledTaskTrigger -Daily -At $Time
    $reportSettings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
        -RestartCount 2 -RestartInterval (New-TimeSpan -Minutes 10)
    Register-ScheduledTask -TaskName $reportTask -Action $action -Trigger $trigger `
        -Settings $reportSettings -Principal $principal -Force `
        -Description "1С «Валовая прибыль организаций - копия» -> лист Регистр -> скриншот -> Telegram" | Out-Null

    Write-Output "Задача '$reportTask' создана: каждый день в $Time."
    Write-Output "Проверить сейчас: Start-ScheduledTask -TaskName $reportTask"
}
