# PostToolUse hook: fires after Edit or Write tool calls.
# Injects a document-module reminder into Claude's context when a code file is modified.
# Deduplicates per session so each file only triggers one reminder.

$inputJson = [Console]::In.ReadToEnd()
try {
    $data = $inputJson | ConvertFrom-Json
} catch {
    exit 0
}

$filePath = $data.tool_input.file_path
if (-not $filePath) { exit 0 }

# Only track recognised code file extensions
$codeExts = @('.py', '.js', '.ts', '.tsx', '.jsx', '.html', '.css', '.sql')
$ext = [System.IO.Path]::GetExtension($filePath).ToLower()
if ($ext -notin $codeExts) { exit 0 }

# Skip files that should never trigger docs
$skipPatterns = @(
    [regex]::Escape('.claude'),
    [regex]::Escape('report_html'),   # served copies, not canonical sources
    '__',                              # temp/debug files
    'node_modules',
    '\.git',
    'settings\.json',
    '\.min\.',
    'package'
)
foreach ($pattern in $skipPatterns) {
    if ($filePath -match $pattern) { exit 0 }
}

# Session dedup: only remind once per file per session
$sessionFile = Join-Path $PSScriptRoot ".doc-reminded"
$reminded = @()
if (Test-Path $sessionFile) {
    $reminded = Get-Content $sessionFile -ErrorAction SilentlyContinue
}
if ($reminded -contains $filePath) { exit 0 }

# Record this file as reminded
Add-Content -Path $sessionFile -Value $filePath -Encoding UTF8

# Output — this goes back into Claude's live context window
$fileName = [System.IO.Path]::GetFileName($filePath)
Write-Output ""
Write-Output "[MODULE-DOC REQUIRED] Code file modified: $fileName"
Write-Output "Per CLAUDE.md §10, you MUST invoke the 'document-module' skill for this module before"
Write-Output "the task is considered complete. Use: Skill('document-module') now."
Write-Output ""
