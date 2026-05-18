# Stop hook: fires just before Claude finishes a turn.
# If any code files were modified this session without confirmed doc generation,
# outputs a visible warning to the user and cleans up the session tracker.

$sessionFile = Join-Path $PSScriptRoot ".doc-reminded"

if (-not (Test-Path $sessionFile)) { exit 0 }

$files = Get-Content $sessionFile -ErrorAction SilentlyContinue | Where-Object { $_ -match '\S' }

if ($files -and $files.Count -gt 0) {
    Write-Output ""
    Write-Output "================================================================"
    Write-Output "  MODULE DOCUMENTATION CHECK"
    Write-Output "================================================================"
    Write-Output "  The following code files were modified this session."
    Write-Output "  Confirm the document-module skill was invoked for each:"
    Write-Output ""
    foreach ($f in $files) {
        $name = [System.IO.Path]::GetFileName($f)
        Write-Output "    - $name"
    }
    Write-Output ""
    Write-Output "  If docs were NOT updated, run: Skill('document-module')"
    Write-Output "  Task is NOT complete without module .md files updated (CLAUDE.md §10)."
    Write-Output "================================================================"
    Write-Output ""
}

# Clean up for next session
Remove-Item $sessionFile -Force -ErrorAction SilentlyContinue
exit 0
