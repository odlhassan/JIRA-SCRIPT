@echo off
REM Restart the report server

setlocal enabledelayedexpansion

cd /d E:\JIRA SCRIPT

echo [restart] Killing any existing Python processes...
REM Try to kill processes by PID from marker files
for /f "tokens=*" %%A in ('dir /b .codex_run_server_*_pid 2^>nul') do (
    for /f "tokens=*" %%B in ('type "%%A"') do (
        echo [restart] Killing PID %%B from file %%A
        taskkill /PID %%B /F 2>nul
    )
    del /f /q "%%A" 2>nul
)

timeout /t 2 /nobreak

echo [restart] Starting fresh server...
echo [restart] Using PORT from environment or default 3000
python run_server.py --no-sync

pause
