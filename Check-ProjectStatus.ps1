# Project Status Check Script
# Verifies that all components are ready for production

param(
    [Parameter(Mandatory=$false)]
    [switch]$Detailed
)

Write-Host "URL-to-Text Transport Agent - Project Status Check" -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green

$issues = @()
$warnings = @()
$successes = @()

# Check project files
Write-Host "`n📁 Checking Project Structure..." -ForegroundColor Yellow

$requiredFiles = @{
    "UrlToTextTransportAgent.csproj" = "Production project file"
    "UrlToTextTransportAgent.Mock.csproj" = "Mock project file" 
    "UrlToTextTransportAgent.Tests.csproj" = "Test project file"
    "UrlToTextTransportAgent.sln" = "Solution file"
    "UrlToTextAgent.cs" = "Main agent implementation"
    "UrlToTextAgentFactory.cs" = "Agent factory"
    "Mock\MockExchangeTypes.cs" = "Mock Exchange API"
    "Tests\UrlToTextAgentTests.cs" = "Unit tests"
    "Tests\TestHelpers.cs" = "Test helpers"
    "App.config" = "Configuration file"
    "README.md" = "Documentation"
}

foreach ($file in $requiredFiles.Keys) {
    if (Test-Path $file) {
        $successes += "✅ $($requiredFiles[$file]): $file"
    } else {
        $issues += "❌ Missing: $file"
    }
}

# Check PowerShell scripts
$scripts = @{
    "Build-Agent.ps1" = "Production build script"
    "Build-Mock.ps1" = "Mock build script"
    "Install-Agent.ps1" = "Installation script"
    "Uninstall-Agent.ps1" = "Uninstallation script"
    "Deploy-Agent.ps1" = "Deployment script"
    "Test-Agent.ps1" = "Testing script"
}

foreach ($script in $scripts.Keys) {
    if (Test-Path $script) {
        $successes += "✅ $($scripts[$script]): $script"
        
        # Check script syntax
        try {
            $null = [System.Management.Automation.PSParser]::Tokenize((Get-Content $script -Raw), [ref]$null)
        } catch {
            $issues += "❌ Syntax error in $script"
        }
    } else {
        $issues += "❌ Missing script: $script"
    }
}

# Check GitHub Actions
Write-Host "`n🔧 Checking GitHub Actions..." -ForegroundColor Yellow

$workflows = @{
    ".github\workflows\ci-cd.yml" = "CI/CD workflow"
    ".github\workflows\test.yml" = "Test workflow"
    ".github\workflows\powershell-test.yml" = "PowerShell validation"
    ".github\workflows\release.yml" = "Release workflow"
}

foreach ($workflow in $workflows.Keys) {
    if (Test-Path $workflow) {
        $successes += "✅ $($workflows[$workflow]): $workflow"
    } else {
        $issues += "❌ Missing workflow: $workflow"
    }
}

# Check build capability
Write-Host "`n🔨 Checking Build Environment..." -ForegroundColor Yellow

# Check for MSBuild
$msbuildPaths = @(
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\*\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\*\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\*\MSBuild\15.0\Bin\MSBuild.exe"
)

$msbuildFound = $false
foreach ($pattern in $msbuildPaths) {
    $paths = Get-ChildItem $pattern -ErrorAction SilentlyContinue
    if ($paths) {
        $successes += "✅ MSBuild found: $($paths[0].FullName)"
        $msbuildFound = $true
        break
    }
}

if (-not $msbuildFound) {
    $warnings += "⚠️ MSBuild not found - install Visual Studio or Build Tools"
}

# Test build mock version
Write-Host "`n🚀 Testing Mock Build..." -ForegroundColor Yellow
try {
    if ($msbuildFound) {
        $result = & ".\Build-Mock.ps1" -Configuration Release 2>&1
        if ($LASTEXITCODE -eq 0) {
            $successes += "✅ Mock build successful"
            
            # Check output files
            if (Test-Path "bin\Release\UrlToTextTransportAgent.Mock.dll") {
                $successes += "✅ Mock DLL created"
                
                # Test DLL loading
                try {
                    Add-Type -Path "bin\Release\UrlToTextTransportAgent.Mock.dll"
                    $factory = New-Object UrlToTextTransportAgent.UrlToTextAgentFactory
                    $successes += "✅ Mock DLL loads and creates agents"
                } catch {
                    $issues += "❌ Mock DLL load test failed: $($_.Exception.Message)"
                }
            } else {
                $issues += "❌ Mock DLL not created"
            }
        } else {
            $issues += "❌ Mock build failed"
        }
    } else {
        $warnings += "⚠️ Skipping build test - MSBuild not available"
    }
} catch {
    $issues += "❌ Build test error: $($_.Exception.Message)"
}

# Check configuration
Write-Host "`n⚙️ Checking Configuration..." -ForegroundColor Yellow

try {
    [xml]$config = Get-Content "App.config"
    $successes += "✅ App.config is valid XML"
    
    # Check critical settings
    $criticalSettings = @("InternalDomains", "LogFilePath", "SkipSignedEmails", "SkipEncryptedEmails")
    foreach ($setting in $criticalSettings) {
        $node = $config.configuration.appSettings.add | Where-Object { $_.key -eq $setting }
        if ($node) {
            $successes += "✅ Configuration setting: $setting = $($node.value)"
        } else {
            $issues += "❌ Missing configuration: $setting"
        }
    }
} catch {
    $issues += "❌ App.config validation failed: $($_.Exception.Message)"
}

# Check documentation
Write-Host "`n📚 Checking Documentation..." -ForegroundColor Yellow

$readmeContent = Get-Content "README.md" -Raw
$requiredSections = @("Features", "Installation", "Configuration", "Security Features", "CI/CD")

foreach ($section in $requiredSections) {
    if ($readmeContent -like "*## $section*") {
        $successes += "✅ README section: $section"
    } else {
        $warnings += "⚠️ README missing section: $section"
    }
}

# Summary
Write-Host "`n" + "=" * 60 -ForegroundColor Green
Write-Host "PROJECT STATUS SUMMARY" -ForegroundColor Green
Write-Host "=" * 60 -ForegroundColor Green

Write-Host "`n✅ SUCCESSES ($($successes.Count)):" -ForegroundColor Green
if ($Detailed -or $successes.Count -le 10) {
    foreach ($success in $successes) {
        Write-Host "   $success" -ForegroundColor Green
    }
} else {
    Write-Host "   $($successes.Count) items successful (use -Detailed to see all)" -ForegroundColor Green
}

if ($warnings.Count -gt 0) {
    Write-Host "`n⚠️ WARNINGS ($($warnings.Count)):" -ForegroundColor Yellow
    foreach ($warning in $warnings) {
        Write-Host "   $warning" -ForegroundColor Yellow
    }
}

if ($issues.Count -gt 0) {
    Write-Host "`n❌ ISSUES ($($issues.Count)):" -ForegroundColor Red
    foreach ($issue in $issues) {
        Write-Host "   $issue" -ForegroundColor Red
    }
    
    Write-Host "`n🔧 NEXT STEPS:" -ForegroundColor Cyan
    Write-Host "   1. Fix the issues listed above" -ForegroundColor Gray
    Write-Host "   2. Re-run this status check" -ForegroundColor Gray
    Write-Host "   3. Proceed with deployment when all issues are resolved" -ForegroundColor Gray
    
    exit 1
} else {
    Write-Host "`n🎉 PROJECT STATUS: READY FOR PRODUCTION!" -ForegroundColor Green
    
    if ($warnings.Count -gt 0) {
        Write-Host "   Note: Address warnings for optimal setup" -ForegroundColor Yellow
    }
    
    Write-Host "`n🚀 DEPLOYMENT OPTIONS:" -ForegroundColor Cyan
    Write-Host "   • Development: .\Deploy-Agent.ps1 -AgentDllPath 'bin\Release\UrlToTextTransportAgent.Mock.dll'" -ForegroundColor Gray
    Write-Host "   • Production: Build on Exchange server, then .\Deploy-Agent.ps1" -ForegroundColor Gray
    Write-Host "   • Testing: .\Test-Agent.ps1 -TestemailTo 'test@domain.com'" -ForegroundColor Gray
}