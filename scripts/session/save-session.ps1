param(
    [Parameter(Mandatory = $true)]
    [string]$Title,

    [Parameter(Mandatory = $true)]
    [string]$UserRequest,

    [Parameter(Mandatory = $true)]
    [string]$WorkDone,

    [string]$NextSteps = "",

    [string]$Notes = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$sessionRoot = Join-Path $repoRoot ".codex-session"
$historyDir = Join-Path $sessionRoot "history"
$briefPath = Join-Path $sessionRoot "session-brief.local.md"

New-Item -ItemType Directory -Path $sessionRoot -Force | Out-Null
New-Item -ItemType Directory -Path $historyDir -Force | Out-Null

if (-not (Test-Path $briefPath)) {
    Set-Content -LiteralPath $briefPath -Value "# Session Brief`r`n" -Encoding UTF8
}

$timestamp = Get-Date
$fileName = $timestamp.ToString("yyyyMMdd-HHmmss") + ".md"
$filePath = Join-Path $historyDir $fileName

$lines = @(
    ("# " + $Title),
    "",
    ("- Time: " + $timestamp.ToString("yyyy-MM-dd HH:mm:ss")),
    ("- User Request: " + $UserRequest),
    "",
    "## Work Done",
    $WorkDone,
    "",
    "## Next Steps",
    $NextSteps,
    "",
    "## Notes",
    $Notes
)

$content = [string]::Join([Environment]::NewLine, $lines)

Set-Content -LiteralPath $filePath -Value $content -Encoding UTF8

Write-Output ("Saved session history to: " + $filePath)
Write-Output "Remember to refresh .codex-session/session-brief.local.md if the active context changed."
