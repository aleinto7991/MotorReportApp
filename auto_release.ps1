param(
    [switch]$DryRun = $false,
    [switch]$NoPush = $false
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host " AUTOMATED RELEASE PIPELINE" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path ".git")) {
    Write-Host " Error: Not a git repository" -ForegroundColor Red
    exit 1
}

Write-Host " Checking for changes..." -ForegroundColor Yellow
$status = git status --porcelain

if (-not $status) {
    Write-Host " No changes to commit" -ForegroundColor Green
    Write-Host ""
    Write-Host " Tip: Make some changes first, then run this script" -ForegroundColor Cyan
    exit 0
}

Write-Host " Changes detected:" -ForegroundColor Green
git status --short
Write-Host ""

if (-not $DryRun) {
    Write-Host " Staging all changes..." -ForegroundColor Yellow
    git add -A
    Write-Host " Changes staged" -ForegroundColor Green
}
else {
    Write-Host " [DRY RUN] Would stage all changes" -ForegroundColor Yellow
}
Write-Host ""

Write-Host " Generating commit message..." -ForegroundColor Yellow
if ($DryRun) {
    Write-Host "[DRY RUN] Commit preview:" -ForegroundColor Yellow
    python auto_commit.py --dry-run
}
else {
    python auto_commit.py
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host " Commit cancelled or failed" -ForegroundColor Red
        exit 1
    }
}
Write-Host ""

Write-Host " Analyzing commits for version bump..." -ForegroundColor Yellow
if ($DryRun) {
    python auto_version.py --dry-run
}
else {
    python auto_version.py
}

if ($LASTEXITCODE -ne 0) {
    Write-Host "  Version bump not needed or failed" -ForegroundColor Yellow
}
Write-Host ""

if (-not $DryRun -and -not $NoPush) {
    Write-Host " Pushing to GitHub..." -ForegroundColor Yellow
    
    $pushResult = git push 2>&1
    $tagsResult = git push --tags 2>&1
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host " Pushed to GitHub successfully" -ForegroundColor Green
    }
    else {
        Write-Host "  Push failed" -ForegroundColor Yellow
        Write-Host " Run manually: git push && git push --tags" -ForegroundColor Cyan
    }
}
elseif ($NoPush) {
    Write-Host "  Skipping push" -ForegroundColor Yellow
}
else {
    Write-Host " [DRY RUN] Would push to GitHub" -ForegroundColor Yellow
}
Write-Host ""

Write-Host "================================" -ForegroundColor Cyan
if ($DryRun) {
    Write-Host " DRY RUN COMPLETE" -ForegroundColor Yellow
    Write-Host "No changes were made" -ForegroundColor Yellow
}
else {
    Write-Host " RELEASE COMPLETE!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "1. Check GitHub for new release" -ForegroundColor White
    Write-Host "2. Build executable: pyinstaller MotorReportApp.spec" -ForegroundColor White
    Write-Host "3. Upload .exe to GitHub release" -ForegroundColor White
}
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""
