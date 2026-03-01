# Build Mock DLL Script
# This script builds the mock version for testing without Exchange dependencies

param(
    [Parameter(Mandatory=$false)]
    [string]$Configuration = "Release",
    
    [Parameter(Mandatory=$false)]
    [switch]$RunTests,
    
    [Parameter(Mandatory=$false)]
    [switch]$Clean
)

Write-Host "URL-to-Text Transport Agent - Mock Build Script" -ForegroundColor Green
Write-Host "===============================================" -ForegroundColor Green

# Clean if requested
if ($Clean) {
    Write-Host "`nCleaning previous builds..." -ForegroundColor Yellow
    if (Test-Path "bin") { Remove-Item "bin" -Recurse -Force }
    if (Test-Path "obj") { Remove-Item "obj" -Recurse -Force }
    if (Test-Path "TestResults") { Remove-Item "TestResults" -Recurse -Force }
    Write-Host "✓ Clean completed" -ForegroundColor Green
}

# Find MSBuild
$msbuildPaths = @(
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\Enterprise\MSBuild\15.0\Bin\MSBuild.exe"
)

$msbuild = $null
foreach ($path in $msbuildPaths) {
    if (Test-Path $path) {
        $msbuild = $path
        break
    }
}

if (-not $msbuild) {
    Write-Host "❌ MSBuild not found. Please install Visual Studio." -ForegroundColor Red
    exit 1
}

Write-Host "`nUsing MSBuild: $msbuild" -ForegroundColor Yellow
Write-Host "Configuration: $Configuration" -ForegroundColor Yellow

try {
    # Build Mock Agent
    Write-Host "`nBuilding Mock Agent..." -ForegroundColor Yellow
    & $msbuild "UrlToTextTransportAgent.Mock.csproj" /p:Configuration=$Configuration /p:Platform="Any CPU" /verbosity:minimal /nologo
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ Mock Agent built successfully!" -ForegroundColor Green
        
        $mockDll = "bin\$Configuration\UrlToTextTransportAgent.Mock.dll"
        if (Test-Path $mockDll) {
            $dllInfo = Get-Item $mockDll
            Write-Host "   DLL: $mockDll" -ForegroundColor Cyan
            Write-Host "   Size: $([math]::Round($dllInfo.Length / 1KB, 2)) KB" -ForegroundColor Cyan
        }
    } else {
        Write-Host "❌ Mock Agent build failed" -ForegroundColor Red
        exit 1
    }
    
    # Build tests if requested
    if ($RunTests) {
        Write-Host "`nBuilding and running tests..." -ForegroundColor Yellow
        
        # Restore NuGet packages for tests
        if (Get-Command nuget -ErrorAction SilentlyContinue) {
            nuget restore UrlToTextTransportAgent.Tests.csproj | Out-Null
        }
        
        # Build test project
        & $msbuild "UrlToTextTransportAgent.Tests.csproj" /p:Configuration=$Configuration /p:Platform="Any CPU" /verbosity:minimal /nologo
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Tests built successfully!" -ForegroundColor Green
            
            # Run tests
            $testDll = "bin\$Configuration\UrlToTextTransportAgent.Tests.dll"
            if (Test-Path $testDll) {
                Write-Host "   Running unit tests..." -ForegroundColor Cyan
                
                # Try to find VSTest
                $vstest = $null
                $vstestPaths = @(
                    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe",
                    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe",
                    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Professional\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe",
                    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe"
                )
                
                foreach ($path in $vstestPaths) {
                    if (Test-Path $path) {
                        $vstest = $path
                        break
                    }
                }
                
                if ($vstest) {
                    & $vstest $testDll /logger:console /nologo
                    
                    if ($LASTEXITCODE -eq 0) {
                        Write-Host "✓ All tests passed!" -ForegroundColor Green
                    } else {
                        Write-Host "❌ Some tests failed" -ForegroundColor Red
                        exit 1
                    }
                } else {
                    Write-Host "⚠️ VSTest not found, skipping test execution" -ForegroundColor Yellow
                    Write-Host "   Tests built successfully but not executed" -ForegroundColor Cyan
                }
            }
        } else {
            Write-Host "❌ Test build failed" -ForegroundColor Red
            exit 1
        }
    }
    
    # Test Mock DLL functionality
    Write-Host "`nTesting Mock DLL functionality..." -ForegroundColor Yellow
    try {
        Add-Type -Path $mockDll
        
        # Test factory creation
        $factory = New-Object UrlToTextTransportAgent.UrlToTextAgentFactory
        $mockServer = New-Object Microsoft.Exchange.Data.Transport.SmtpServer
        $agent = $factory.CreateAgent($mockServer)
        
        if ($agent) {
            Write-Host "✓ Mock Agent factory working correctly" -ForegroundColor Green
        } else {
            Write-Host "❌ Agent creation failed" -ForegroundColor Red
            exit 1
        }
    } catch {
        Write-Host "❌ Mock DLL test failed: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "`n" + "=" * 50 -ForegroundColor Green
    Write-Host "MOCK BUILD COMPLETED SUCCESSFULLY!" -ForegroundColor Green
    Write-Host "=" * 50 -ForegroundColor Green
    
    Write-Host "`nOutput Files:" -ForegroundColor Cyan
    Write-Host "  Mock DLL: bin\$Configuration\UrlToTextTransportAgent.Mock.dll" -ForegroundColor Gray
    Write-Host "  Config: bin\$Configuration\UrlToTextTransportAgent.Mock.dll.config" -ForegroundColor Gray
    
    if ($RunTests) {
        Write-Host "  Test DLL: bin\$Configuration\UrlToTextTransportAgent.Tests.dll" -ForegroundColor Gray
    }
    
    Write-Host "`nNext Steps:" -ForegroundColor Cyan
    Write-Host "  • Use the Mock DLL for testing without Exchange dependencies" -ForegroundColor Gray
    Write-Host "  • Build the production version on an Exchange server" -ForegroundColor Gray
    Write-Host "  • Run .\Deploy-Agent.ps1 for automated deployment" -ForegroundColor Gray
    
} catch {
    Write-Host "❌ Build failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}