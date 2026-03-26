<#
.SYNOPSIS
    Update DrawingAI Pro on the server to the latest version from GitHub.
    Run this every time you push a new version.

.DESCRIPTION
    1. Stops the service (if running)
    2. Pulls latest code from GitHub
    3. Updates Python dependencies (if changed)
    4. Restarts the service

.EXAMPLE
    .\update.ps1
    .\update.ps1 -Branch feature/new-stage
    .\update.ps1 -SkipRestart
#>

param(
    [string]$Branch = "main",
    [switch]$SkipRestart,
    [switch]$ForceRequirements
)

$ErrorActionPreference = "Stop"
$installDir = Split-Path -Parent $PSScriptRoot  # assume script is in deploy\

Write-Host "`n🔄 DrawingAI Pro — Update" -ForegroundColor Cyan
Write-Host "  Directory: $installDir" -ForegroundColor Gray
Write-Host "  Branch: $Branch`n" -ForegroundColor Gray

Set-Location $installDir

# ── 1. Stop service ──────────────────────────────────────────────────
$serviceName = "DrawingAI"
$serviceRunning = $false

try {
    $svc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    if ($svc -and $svc.Status -eq "Running") {
        Write-Host "[1/5] Stopping service..." -ForegroundColor Yellow
        Stop-Service $serviceName -Force
        Start-Sleep -Seconds 3
        $serviceRunning = $true
        Write-Host "  ✓ Service stopped" -ForegroundColor Green
    } else {
        Write-Host "[1/5] Service not running — skipping" -ForegroundColor Gray
    }
} catch {
    Write-Host "[1/5] No service found (running in GUI mode?)" -ForegroundColor Gray
}

# ── 2. Backup current config ─────────────────────────────────────────
Write-Host "[2/5] Backing up config files..." -ForegroundColor Yellow

$backupDir = "$installDir\deploy\backups\$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

$configFiles = @(".env", "email_config.json", "automation_config.json", "automation_state.json")
foreach ($f in $configFiles) {
    if (Test-Path "$installDir\$f") {
        Copy-Item "$installDir\$f" "$backupDir\$f"
    }
}
Write-Host "  ✓ Backup saved to $backupDir" -ForegroundColor Green

# ── 3. Pull latest code ──────────────────────────────────────────────
Write-Host "[3/5] Pulling latest from GitHub ($Branch)..." -ForegroundColor Yellow

# Save current commit for comparison
$oldCommit = git rev-parse HEAD 2>$null

git fetch origin
git checkout $Branch
git pull origin $Branch

$newCommit = git rev-parse HEAD
if ($oldCommit -eq $newCommit) {
    Write-Host "  ✓ Already up to date ($($newCommit.Substring(0,7)))" -ForegroundColor Green
} else {
    Write-Host "  ✓ Updated: $($oldCommit.Substring(0,7)) → $($newCommit.Substring(0,7))" -ForegroundColor Green
    # Show what changed
    git log --oneline "$oldCommit..$newCommit" | ForEach-Object {
        Write-Host "    $_" -ForegroundColor DarkGray
    }
}

# ── 4. Update dependencies ───────────────────────────────────────────
Write-Host "[4/5] Checking dependencies..." -ForegroundColor Yellow

# Check if requirements.txt changed
$reqChanged = git diff "$oldCommit" "$newCommit" -- requirements.txt 2>$null
if ($reqChanged -or $ForceRequirements) {
    Write-Host "  📦 requirements.txt changed — installing..." -ForegroundColor Yellow
    & .\.venv\Scripts\pip.exe install -r requirements.txt --quiet
    Write-Host "  ✓ Dependencies updated" -ForegroundColor Green
} else {
    Write-Host "  ✓ No dependency changes" -ForegroundColor Green
}

# ── 5. Restart service ───────────────────────────────────────────────
if ($SkipRestart) {
    Write-Host "[5/5] Restart skipped (--SkipRestart)" -ForegroundColor Gray
} elseif ($serviceRunning) {
    Write-Host "[5/5] Starting service..." -ForegroundColor Yellow
    Start-Service $serviceName
    Start-Sleep -Seconds 2
    $svc = Get-Service -Name $serviceName
    Write-Host "  ✓ Service status: $($svc.Status)" -ForegroundColor Green
} else {
    Write-Host "[5/5] Service was not running — start manually or via GUI" -ForegroundColor Gray
}

# ── Summary ───────────────────────────────────────────────────────────
Write-Host "`n✅ Update complete!" -ForegroundColor Green
Write-Host "  Version: $(git log --oneline -1)" -ForegroundColor White
Write-Host "  Backup: $backupDir" -ForegroundColor Gray
Write-Host ""
