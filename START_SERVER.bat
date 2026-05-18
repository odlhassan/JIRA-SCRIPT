@echo off
REM =====================================================
REM Report Server Restart Script
REM =====================================================
setlocal enabledelayedexpansion

cd /d "E:\JIRA SCRIPT"

echo.
echo =====================================================
echo Report Server Restart Sequence
echo =====================================================
echo.

REM Step 1: Kill processes from PID marker files
echo [step 1] Killing old processes from marker files...
set pidlist=
for /f "tokens=*" %%F in ('dir /b .codex_run_server_*_pid 2^>nul') do (
    for /f "tokens=*" %%P in ('type "%%F"') do (
        echo [kill] Found PID %%P in %%F
        taskkill /PID %%P /F 2>nul
        if !errorlevel! equ 0 (
            echo [kill] Successfully killed PID %%P
        ) else (
            echo [kill] PID %%P not found or already terminated
        )
    )
)

echo.
echo [step 2] Cleaning up marker files...
for /f "tokens=*" %%F in ('dir /b .codex_run_server_*_pid 2^>nul') do (
    del /f /q "%%F" 2>nul
    echo [clean] Deleted %%F
)

timeout /t 2 /nobreak

echo.
echo [step 3] Starting fresh server...
echo [server] Command: python run_server.py --no-sync
echo.

python run_server.py --no-sync

echo.
echo Server stopped.
pause
