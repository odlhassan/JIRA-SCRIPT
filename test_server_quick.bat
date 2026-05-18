@echo off
REM Quick server test - just test if already running
cd /d "E:\JIRA SCRIPT"

echo Testing existing server...
curl -s "http://127.0.0.1:3000/api/team-capacity-planner/work-items?project=MN" > nul 2>&1

if %errorlevel% equ 0 (
    echo Server is already running on 3000
    echo Testing work-items endpoint...
    curl -s "http://127.0.0.1:3000/api/team-capacity-planner/work-items?project=MN" | findstr "MN-137"
    if %errorlevel% equ 0 (
        echo SUCCESS: MN-137 found
    ) else (
        echo Server running but MN-137 not found
    )
) else (
    echo No server on port 3000, trying other ports...
    for %%P in (3001 3002 3003 3004 3005) do (
        curl -s "http://127.0.0.1:%%P/api/team-capacity-planner/work-items?project=MN" > nul 2>&1
        if !errorlevel! equ 0 (
            echo Server found on port %%P
            curl -s "http://127.0.0.1:%%P/api/team-capacity-planner/work-items?project=MN" | findstr "MN-137"
            if !errorlevel! equ 0 (
                echo SUCCESS: MN-137 found
            )
            exit /b 0
        )
    )
    echo No server responding, needs restart
)
