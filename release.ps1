# Release helper wrapper
# Usage: .\release.ps1 [-DryRun]

param(
    [switch]$Preview,
    [switch]$DryRun,
    [switch]$Release
)

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "Error: virtualenv not found (.venv). Create it or run from environment." -ForegroundColor Red
    exit 1
}

& .\.venv\Scripts\Activate.ps1

$python = ".\.venv\Scripts\python.exe"
$args = @()

if ($Preview) { $args += "--preview" }
elseif ($DryRun) { $args += "--dry-run" }
elseif ($Release) { $args += "--release" }
else {
    Write-Host "No mode provided. Choose: preview / dry-run / release"
    $choice = Read-Host "Enter 'preview', 'dry-run' or 'release' (default preview)"
    if ($choice -eq 'release') { $args += "--release" }
    elseif ($choice -eq 'dry-run') { $args += "--dry-run" }
    else { $args += "--preview" }
}

# If performing a real release, stage and commit all local changes first
if ($args -contains "--release") {
    $status = git status --porcelain
    if ($status) {
        Write-Host "Staging and committing local changes before release..."
        git add -A
        # Use a clear commit message; if no changes to commit, git will exit non-zero
        try {
            git commit -m "chore(release): pre-release auto-commit"
        } catch {
            Write-Host "No changes were committed (maybe nothing to commit). Continuing."
        }
    } else {
        Write-Host "Working tree clean. Proceeding to release."
    }
}

& $python release.py @args
