# Complete deployment script for URL-to-Text Transport Agent
# This script builds and installs the agent in one step

param(
    [Parameter(Mandatory=$false)]
    [string]$Configuration = "Release",
    
    [Parameter(Mandatory=$false)]
    [string]$AgentName = "UrlToTextAgent",
    
    [Parameter(Mandatory=$false)]
    [string]$InstallPath = "$env:ProgramFiles\Exchange\TransportAgents",
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipBuild,
    
    [Parameter(Mandatory=$false)]
    [switch]$TestAfterInstall
)

$ErrorActionPreference = "Stop"

Write-Host "URL-to-Text Transport Agent - Complete Deployment" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green

# Step 1: Build the agent (unless skipped)
if (!$SkipBuild) {
    Write-Host "`n[1/4] Building the Transport Agent..." -ForegroundColor Yellow
    try {
        & .\Build-Agent.ps1 -Configuration $Configuration
        
        $dllPath = ".\bin\$Configuration\UrlToTextTransportAgent.dll"
        if (!(Test-Path $dllPath)) {
            throw "Build completed but DLL not found at: $dllPath"
        }
        
        Write-Host "✓ Build completed successfully" -ForegroundColor Green
    } catch {
        Write-Host "✗ Build failed: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "`n[1/4] Skipping build (as requested)..." -ForegroundColor Yellow
    $dllPath = ".\bin\$Configuration\UrlToTextTransportAgent.dll"
    
    if (!(Test-Path $dllPath)) {
        Write-Host "✗ DLL not found at: $dllPath" -ForegroundColor Red
        Write-Host "Remove -SkipBuild parameter to build the project first" -ForegroundColor Yellow
        exit 1
    }
}

# Step 2: Prepare installation directory
Write-Host "`n[2/4] Preparing installation directory..." -ForegroundColor Yellow
try {
    if (!(Test-Path $InstallPath)) {
        New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
        Write-Host "✓ Created directory: $InstallPath" -ForegroundColor Green
    } else {
        Write-Host "✓ Directory exists: $InstallPath" -ForegroundColor Green
    }
    
    # Copy DLL and config to installation directory
    $targetDll = Join-Path $InstallPath "UrlToTextTransportAgent.dll"
    $targetConfig = Join-Path $InstallPath "UrlToTextTransportAgent.dll.config"
    
    Copy-Item $dllPath $targetDll -Force
    Write-Host "✓ Copied DLL to: $targetDll" -ForegroundColor Green
    
    $configPath = ".\bin\$Configuration\UrlToTextTransportAgent.dll.config"
    if (Test-Path $configPath) {
        Copy-Item $configPath $targetConfig -Force
        Write-Host "✓ Copied config to: $targetConfig" -ForegroundColor Green
    }
    
} catch {
    Write-Host "✗ Error preparing installation: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Step 3: Install the agent
Write-Host "`n[3/4] Installing the Transport Agent..." -ForegroundColor Yellow
try {
    & .\Install-Agent.ps1 -AgentDllPath $targetDll -AgentName $AgentName
    Write-Host "✓ Agent installed successfully" -ForegroundColor Green
} catch {
    Write-Host "✗ Installation failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Step 4: Test the installation (if requested)
if ($TestAfterInstall) {
    Write-Host "`n[4/4] Testing the installation..." -ForegroundColor Yellow
    try {
        & .\Test-Agent.ps1 -AgentName $AgentName
        Write-Host "✓ Testing completed" -ForegroundColor Green
    } catch {
        Write-Host "⚠ Testing encountered issues: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "`n[4/4] Skipping tests (use -TestAfterInstall to enable)" -ForegroundColor Yellow
}

# Summary
# Get default log path
$defaultLogPath = "$env:ProgramData\Microsoft\Exchange\Logs\UrlToTextAgent.log"

Write-Host "`n" + "=" * 50 -ForegroundColor Green
Write-Host "DEPLOYMENT COMPLETED SUCCESSFULLY!" -ForegroundColor Green
Write-Host "=" * 50 -ForegroundColor Green

Write-Host "`nAgent Details:" -ForegroundColor Cyan
Write-Host "  Name: $AgentName" -ForegroundColor Gray
Write-Host "  DLL Location: $targetDll" -ForegroundColor Gray
Write-Host "  Log File: $defaultLogPath" -ForegroundColor Gray

Write-Host "`nNext Steps:" -ForegroundColor Cyan
Write-Host "  1. Send test emails from external accounts with URLs" -ForegroundColor Gray
Write-Host "  2. Verify URLs are converted to plain text (no longer clickable)" -ForegroundColor Gray
Write-Host "  3. Monitor the log file for processing activity" -ForegroundColor Gray
Write-Host "  4. Run .\Test-Agent.ps1 for detailed status" -ForegroundColor Gray

Write-Host "`nUseful Commands:" -ForegroundColor Cyan
Write-Host "  Get-TransportAgent -Identity '$AgentName'" -ForegroundColor Gray
Write-Host "  Get-Content '$defaultLogPath' -Tail 10" -ForegroundColor Gray
Write-Host "  .\Uninstall-Agent.ps1  # To remove the agent" -ForegroundColor Gray

Write-Host "`nFor support, see README.md for troubleshooting steps." -ForegroundColor Yellow