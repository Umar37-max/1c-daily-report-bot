<#
    Чтобы ежедневный отчёт работал после выхода из RDP.

    Прогон водит 1С по окнам, а окна живут на рабочем столе сессии. Пока по
    RDP подключён человек — стол есть. Закрыли окно RDP крестиком — сессия
    продолжает работать (бот отвечает, расписание срабатывает), но рабочего
    стола у неё больше нет: нажатия по 1С «проходят без видимого эффекта»,
    и прогон падает с ошибкой

        Could not click '...' without a cursor - tried: ...
        call succeeded but had no visible effect

    Лечение: выходить из RDP не крестиком, а переводом сессии на консоль
    сервера — `tscon <id> /dest:console`. Сессия остаётся живой и с полноценным
    рабочим столом.

    Скрипт делает три вещи:
      1. кладёт на рабочий стол ярлык «Отключиться (оставить рабочим)»;
      2. выключает заставку и блокировку экрана для этого пользователя
         (заблокированный экран — это другой рабочий стол, 1С на нём тоже
         не нажимается);
      3. показывает текущее состояние сессии.

    Запускать ОДИН раз, на сервере, от администратора:
        powershell -ExecutionPolicy Bypass -File C:\daily-report\tools\setup_rdp_console.ps1
#>

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$cmdPath = Join-Path $root "disconnect.cmd"

# ── 1. Команда отключения ──────────────────────────────────────────────
# %SESSIONNAME% — имя текущей сессии (RDP-Tcp#NN); tscon понимает и его.
$cmd = @"
@echo off
rem Выход из RDP с сохранением живого рабочего стола.
rem Сессия переезжает на консоль сервера: 1С и бот продолжают работать.
tscon %SESSIONNAME% /dest:console
"@
Set-Content -Path $cmdPath -Value $cmd -Encoding ASCII
Write-Host "Создан $cmdPath"

# ── 2. Ярлык на рабочем столе ──────────────────────────────────────────
$desktop = [Environment]::GetFolderPath("Desktop")
$linkPath = Join-Path $desktop "Отключиться (оставить рабочим).lnk"
$shell = New-Object -ComObject WScript.Shell
$link = $shell.CreateShortcut($linkPath)
$link.TargetPath = $cmdPath
$link.WorkingDirectory = $root
$link.IconLocation = "shell32.dll,27"
$link.Description = "Выйти из RDP, не убивая рабочий стол (нужно для 1С)"
$link.Save()
Write-Host "Создан ярлык: $linkPath"

# ── 3. Заставка и блокировка экрана — выключить ────────────────────────
$cp = "HKCU:\Control Panel\Desktop"
Set-ItemProperty -Path $cp -Name ScreenSaveActive  -Value "0"
Set-ItemProperty -Path $cp -Name ScreenSaveTimeOut -Value "0"
Set-ItemProperty -Path $cp -Name ScreenSaverIsSecure -Value "0"
Write-Host "Заставка и блокировка экрана выключены для $env:USERNAME"

# ── 4. Что сейчас ──────────────────────────────────────────────────────
Write-Host ""
Write-Host "Состояние сессий:"
query session

Write-Host ""
Write-Host "Проверка глазами прогона:"
$py = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) { $py = "python" }
& $py (Join-Path $root "src\desktop_session.py")

Write-Host ""
Write-Host "Готово. Дальше: уходя с сервера, запускай ярлык"
Write-Host "  «Отключиться (оставить рабочим)»"
Write-Host "вместо крестика окна RDP."
Write-Host "Если всё-таки закрыли крестиком — прогон теперь сам вернёт"
Write-Host "рабочий стол перед запуском 1С (src\desktop_session.py)."
