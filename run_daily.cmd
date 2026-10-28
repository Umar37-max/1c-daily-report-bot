@echo off
rem Ежедневный прогон. Ставится в Планировщик задач (tools\setup_schedule.ps1).
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value 2^>nul ^| find "="') do set DT=%%I
python daily_report.py >> "logs\daily_%DT:~0,8%.log" 2>&1
exit /b %ERRORLEVEL%
