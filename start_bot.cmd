@echo off
REM Runs the Telegram bot IN THE INTERACTIVE USER SESSION (not as a service).
REM A Windows service lives in session 0, which has no user desktop, so the
REM 1C thin client cannot open its window there and every run fails - while
REM the same launch works by hand in an RDP session (verified on the server
REM of the neighbouring project). Register this file in Task Scheduler for
REM "at logon" (tools\setup_schedule.ps1).
REM The loop replaces service auto-restart: if bot.py dies it comes back.
REM ASCII only on purpose: cmd.exe reads .cmd files in the OEM codepage and
REM UTF-8 text here breaks parsing.
REM CRLF line endings on purpose too: with LF-only endings cmd.exe can fail
REM to find the :loop label, and the bot is not restarted after a crash.

cd /d "%~dp0"

set PYTHONIOENCODING=utf-8
set VIRTUAL_INPUT_ALLOW_REAL=0

if not exist logs mkdir logs

:loop
echo [%date% %time%] starting bot.py >> logs\bot.log
python bot.py >> logs\bot.log 2>&1
echo [%date% %time%] bot.py exited, restarting in 10s >> logs\bot.log
timeout /t 10 /nobreak >nul
goto loop
