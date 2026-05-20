@echo off
cd /d "E:\JIRA SCRIPT"

echo ===== GIT STATUS =====
git status

echo.
echo ===== RECENT COMMITS =====
git log --oneline -5

echo.
echo ===== CHECKING BRANCH =====
git branch --show-current

echo.
echo ===== GIT DIFF =====
git diff --stat

if errorlevel 0 (
    echo.
    echo ===== CHECKING FOR UNCOMMITTED CHANGES =====
    git status --porcelain > __git_status_output.txt
    
    REM Count lines in status output
    for /f %%A in ('find /c /v "" ^< __git_status_output.txt') do (
        set "line_count=%%A"
    )
    
    if %line_count% gtr 0 (
        echo Found uncommitted changes. Staging and committing...
        
        git add -A
        git commit -m "chore: sync latest changes" --author "Copilot <223556219+Copilot@users.noreply.github.com>"
        
        echo.
        echo ===== PUSHING TO GITHUB =====
        git push origin main
        
        if errorlevel 0 (
            echo Push successful!
        ) else (
            echo Push failed with error code !errorlevel!
        )
    ) else (
        echo No uncommitted changes found. Checking if push is needed...
        git push origin main
        if errorlevel 0 (
            echo Push completed.
        ) else (
            echo No changes to push or push failed.
        )
    )
    
    del __git_status_output.txt 2>nul
)
