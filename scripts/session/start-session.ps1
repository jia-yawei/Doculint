param(
    [int]$RecentCount = 3
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$sessionRoot = Join-Path $repoRoot ".codex-session"
$briefPath = Join-Path $sessionRoot "session-brief.local.md"
$briefExamplePath = Join-Path $sessionRoot "session-brief.local.example.md"
$historyDir = Join-Path $sessionRoot "history"
$lastReadPath = Join-Path $sessionRoot "last-read.txt"

New-Item -ItemType Directory -Path $sessionRoot -Force | Out-Null
New-Item -ItemType Directory -Path $historyDir -Force | Out-Null

if (-not (Test-Path $briefPath)) {
    if (Test-Path $briefExamplePath) {
        Copy-Item -LiteralPath $briefExamplePath -Destination $briefPath
    } else {
        Set-Content -LiteralPath $briefPath -Value "# Session Brief`r`n" -Encoding UTF8
    }
}

$now = Get-Date
Set-Content -LiteralPath $lastReadPath -Value $now.ToString("yyyy-MM-dd HH:mm:ss") -Encoding UTF8

Write-Output "=== Session Brief ==="
Get-Content -LiteralPath $briefPath -Encoding UTF8

$recentFiles = Get-ChildItem -LiteralPath $historyDir -File |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First $RecentCount

if ($recentFiles.Count -gt 0) {
    Write-Output ""
    Write-Output "=== Recent History ==="
    foreach ($file in $recentFiles) {
        Write-Output ""
        Write-Output ("--- " + $file.Name + " ---")
        Get-Content -LiteralPath $file.FullName -Encoding UTF8
    }
} else {
    Write-Output ""
    Write-Output "=== Recent History ==="
    Write-Output "No session history yet."
}
