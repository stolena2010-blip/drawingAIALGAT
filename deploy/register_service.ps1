<#
.SYNOPSIS
    Register/Unregister DrawingAI as a Windows Service using NSSM.

.EXAMPLE
    .\register_service.ps1 -Action install
    .\register_service.ps1 -Action remove
    .\register_service.ps1 -Action restart
#>
#Requires -RunAsAdministrator

param(
    [ValidateSet("install", "remove", "start", "stop", "restart", "status")]
    [string]$Action = "install"
)

$serviceName = "DrawingAI"
$installDir = Split-Path -Parent $PSScriptRoot
$pythonExe = "$installDir\.venv\Scripts\python.exe"
$mainScript = "main.py"
$nssmPath = "$installDir\deploy\nssm.exe"

# Check NSSM
if (-not (Test-Path $nssmPath)) {
    # Try system PATH
    $nssmPath = (Get-Command nssm -ErrorAction SilentlyContinue).Source
    if (-not $nssmPath) {
        Write-Host "❌ NSSM not found!" -ForegroundColor Red
        Write-Host "Download from https://nssm.cc/download" -ForegroundColor Yellow
        Write-Host "Place nssm.exe in $installDir\deploy\" -ForegroundColor Yellow
        exit 1
    }
}

switch ($Action) {
    "install" {
        Write-Host "Installing service '$serviceName'..." -ForegroundColor Cyan
        
        & $nssmPath install $serviceName $pythonExe $mainScript
        & $nssmPath set $serviceName AppDirectory $installDir
        & $nssmPath set $serviceName DisplayName "DrawingAI Pro Automation"
        & $nssmPath set $serviceName Description "AI-powered engineering drawing extraction and email automation"
        & $nssmPath set $serviceName AppStdout "$installDir\logs\service_stdout.log"
        & $nssmPath set $serviceName AppStderr "$installDir\logs\service_stderr.log"
        & $nssmPath set $serviceName AppRotateFiles 1
        & $nssmPath set $serviceName AppRotateBytes 5000000  # 5MB
        & $nssmPath set $serviceName AppEnvironmentExtra "PYTHONIOENCODING=utf-8"
        & $nssmPath set $serviceName Start SERVICE_DEMAND_START  # Manual start
        
        Write-Host "✓ Service installed (manual start)" -ForegroundColor Green
        Write-Host "  Start: .\register_service.ps1 -Action start" -ForegroundColor Gray
    }
    "remove" {
        Write-Host "Removing service '$serviceName'..." -ForegroundColor Yellow
        & $nssmPath stop $serviceName 2>$null
        & $nssmPath remove $serviceName confirm
        Write-Host "✓ Service removed" -ForegroundColor Green
    }
    "start" {
        & $nssmPath start $serviceName
        Write-Host "✓ Service started" -ForegroundColor Green
    }
    "stop" {
        & $nssmPath stop $serviceName
        Write-Host "✓ Service stopped" -ForegroundColor Green
    }
    "restart" {
        & $nssmPath restart $serviceName
        Write-Host "✓ Service restarted" -ForegroundColor Green
    }
    "status" {
        & $nssmPath status $serviceName
    }
}
