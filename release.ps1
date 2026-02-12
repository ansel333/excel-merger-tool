# Release helper wrapper
# Usage: .\release.ps1 [-DryRun]

param(
    [switch]$DryRun
)

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "Error: virtualenv not found (.venv). Create it or run from environment." -ForegroundColor Red
    exit 1
}

& .\.venv\Scripts\Activate.ps1

$python = ".\.venv\Scripts\python.exe"
$args = @()
if ($DryRun) { $args += "--dry-run" }

& $python release.py @args
