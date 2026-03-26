#Requires -RunAsAdministrator
<#
.SYNOPSIS
    First-time server installation for DrawingAI Pro.
    Run this ONCE on a new Windows Server.

.DESCRIPTION
    1. Clones the GitHub repo
    2. Creates Python venv + installs dependencies
    3. Installs Tesseract OCR
    4. Creates placeholder config files
    5. Registers as a Windows Service (via NSSM)

.EXAMPLE
    # Open PowerShell as Administrator on the server, then:
    .\install_server.ps1 -GitHubRepo "https://github.com/YOUR_USER/AI-DRAW.git"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$GitHubRepo,

    [string]$InstallDir = "C:\DrawingAI",
    [string]$PythonVersion = "3.11"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  DrawingAI Pro — Server Installation" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# ── 1. Prerequisites check ───────────────────────────────────────────
Write-Host "[1/7] Checking prerequisites..." -ForegroundColor Yellow

# Check Git
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "  ❌ Git not found. Install from https://git-scm.com/download/win" -ForegroundColor Red
    exit 1
}
Write-Host "  ✓ Git: $(git --version)" -ForegroundColor Green

# Check Python
$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    Write-Host "  ❌ Python not found. Install Python $PythonVersion from https://python.org" -ForegroundColor Red
    exit 1
}
$pyVer = python --version 2>&1
Write-Host "  ✓ Python: $pyVer" -ForegroundColor Green

# ── 2. Clone repository ──────────────────────────────────────────────
Write-Host "`n[2/7] Cloning repository..." -ForegroundColor Yellow

if (Test-Path $InstallDir) {
    Write-Host "  ⚠ Directory $InstallDir already exists." -ForegroundColor Yellow
    $confirm = Read-Host "  Overwrite? (y/N)"
    if ($confirm -ne "y") {
        Write-Host "  Aborted." -ForegroundColor Red
        exit 1
    }
    Remove-Item $InstallDir -Recurse -Force
}

git clone $GitHubRepo $InstallDir
Set-Location $InstallDir
Write-Host "  ✓ Cloned to $InstallDir" -ForegroundColor Green

# ── 3. Python virtual environment ────────────────────────────────────
Write-Host "`n[3/7] Creating Python virtual environment..." -ForegroundColor Yellow

python -m venv .venv
& .\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet
Write-Host "  ✓ Virtual environment ready" -ForegroundColor Green

# ── 4. Tesseract OCR ─────────────────────────────────────────────────
Write-Host "`n[4/7] Checking Tesseract OCR..." -ForegroundColor Yellow

$tesseractPath = "C:\Program Files\Tesseract-OCR\tesseract.exe"
if (Test-Path $tesseractPath) {
    Write-Host "  ✓ Tesseract found at $tesseractPath" -ForegroundColor Green
} else {
    Write-Host "  ⚠ Tesseract not found!" -ForegroundColor Yellow
    Write-Host "  Download from: https://github.com/UB-Mannheim/tesseract/wiki" -ForegroundColor Yellow
    Write-Host "  Install with Hebrew language pack (heb)" -ForegroundColor Yellow
}

# ── 5. Create config file placeholders ────────────────────────────────
Write-Host "`n[5/7] Creating config files..." -ForegroundColor Yellow

if (-not (Test-Path ".env")) {
    Copy-Item ".env.example" ".env"
    Write-Host "  ✓ Created .env from .env.example — EDIT WITH YOUR KEYS!" -ForegroundColor Yellow
} else {
    Write-Host "  ✓ .env already exists" -ForegroundColor Green
}

if (-not (Test-Path "email_config.json")) {
    if (Test-Path "email_config.example.json") {
        Copy-Item "email_config.example.json" "email_config.json"
        Write-Host "  ✓ Created email_config.json — EDIT WITH YOUR SETTINGS!" -ForegroundColor Yellow
    }
}

# Create working directories
$dirs = @("logs", "TEMP")
foreach ($d in $dirs) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Path $d -Force | Out-Null
    }
}
Write-Host "  ✓ Working directories created" -ForegroundColor Green

# ── 6. Create automation folders ──────────────────────────────────────
Write-Host "`n[6/7] Creating automation folders..." -ForegroundColor Yellow

$automationDirs = @(
    "C:\Users\$env:USERNAME\Desktop\automation\from",
    "C:\Users\$env:USERNAME\Desktop\automation\to",
    "C:\Users\$env:USERNAME\Desktop\automation\archive"
)
foreach ($d in $automationDirs) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Path $d -Force | Out-Null
        Write-Host "  ✓ Created $d" -ForegroundColor Green
    }
}

# ── 7. Install as Windows Service (via NSSM) ─────────────────────────
Write-Host "`n[7/7] Windows Service setup..." -ForegroundColor Yellow

$nssmPath = "$InstallDir\deploy\nssm.exe"
if (-not (Test-Path $nssmPath)) {
    Write-Host "  ⚠ NSSM not found at $nssmPath" -ForegroundColor Yellow
    Write-Host "  Download from: https://nssm.cc/download" -ForegroundColor Yellow
    Write-Host "  Place nssm.exe in $InstallDir\deploy\" -ForegroundColor Yellow
    Write-Host "  Then run: .\deploy\register_service.ps1" -ForegroundColor Yellow
} else {
    & $nssmPath install DrawingAI "$InstallDir\.venv\Scripts\python.exe" "main.py"
    & $nssmPath set DrawingAI AppDirectory $InstallDir
    & $nssmPath set DrawingAI DisplayName "DrawingAI Pro Automation"
    & $nssmPath set DrawingAI Description "AI Drawing extraction and email automation"
    & $nssmPath set DrawingAI AppStdout "$InstallDir\logs\service_stdout.log"
    & $nssmPath set DrawingAI AppStderr "$InstallDir\logs\service_stderr.log"
    & $nssmPath set DrawingAI AppRotateFiles 1
    & $nssmPath set DrawingAI AppRotateBytes 5000000
    & $nssmPath set DrawingAI Start SERVICE_DEMAND_START
    Write-Host "  ✓ Service 'DrawingAI' registered (manual start)" -ForegroundColor Green
    Write-Host "  Start with: nssm start DrawingAI" -ForegroundColor Cyan
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  ✅ Installation complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "`nNext steps:" -ForegroundColor White
Write-Host "  1. Edit $InstallDir\.env with your Azure keys" -ForegroundColor White
Write-Host "  2. Edit $InstallDir\email_config.json" -ForegroundColor White
Write-Host "  3. Edit $InstallDir\automation_config.json" -ForegroundColor White
Write-Host "  4. Test: cd $InstallDir; .\.venv\Scripts\python.exe main.py" -ForegroundColor White
Write-Host "  5. Start service: nssm start DrawingAI" -ForegroundColor White
Write-Host ""
