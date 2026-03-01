# Build script for URL-to-Text Transport Agent
# Run this script to build the transport agent DLL

param(
    [Parameter(Mandatory=$false)]
    [string]$Configuration = "Release",
    
    [Parameter(Mandatory=$false)]
    [string]$MSBuildPath = ""
)

# Find MSBuild if not specified
if ([string]::IsNullOrEmpty($MSBuildPath)) {
    $possiblePaths = @(
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\Enterprise\MSBuild\15.0\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\MSBuild\14.0\Bin\MSBuild.exe"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $MSBuildPath = $path
            break
        }
    }
}

if ([string]::IsNullOrEmpty($MSBuildPath) -or !(Test-Path $MSBuildPath)) {
    Write-Host "MSBuild not found. Please install Visual Studio or specify MSBuildPath parameter." -ForegroundColor Red
    exit 1
}

try {
    Write-Host "Building URL-to-Text Transport Agent..." -ForegroundColor Green
    Write-Host "Using MSBuild: $MSBuildPath" -ForegroundColor Yellow
    Write-Host "Configuration: $Configuration" -ForegroundColor Yellow
    
    # Build the project
    & $MSBuildPath "UrlToTextTransportAgent.csproj" /p:Configuration=$Configuration /p:Platform="Any CPU" /verbosity:minimal
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Build completed successfully!" -ForegroundColor Green
        Write-Host "Output DLL: .\bin\$Configuration\UrlToTextTransportAgent.dll" -ForegroundColor Cyan
        
        if (Test-Path ".\bin\$Configuration\UrlToTextTransportAgent.dll") {
            Write-Host "DLL file created successfully!" -ForegroundColor Green
        }
    } else {
        Write-Host "Build failed with exit code: $LASTEXITCODE" -ForegroundColor Red
        exit 1
    }
    
} catch {
    Write-Host "Error during build: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}