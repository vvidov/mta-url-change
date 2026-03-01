# Exchange Transport Agent Installation Script
# Run this script as Administrator on the Exchange Server

param(
    [Parameter(Mandatory=$true)]
    [string]$AgentDllPath,
    
    [Parameter(Mandatory=$false)]
    [string]$AgentName = "UrlToTextAgent",
    
    [Parameter(Mandatory=$false)]
    [string]$Priority = 1
)

# Check if Exchange Management Shell is loaded
if (!(Get-Command Get-TransportAgent -ErrorAction SilentlyContinue)) {
    Write-Host "Loading Exchange Management Shell..." -ForegroundColor Yellow
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}

try {
    Write-Host "Installing Transport Agent: $AgentName" -ForegroundColor Green
    
    # Install the transport agent
    Install-TransportAgent -Name $AgentName -TransportAgentFactory "UrlToTextTransportAgent.UrlToTextAgentFactory" -AssemblyPath $AgentDllPath
    
    # Set priority
    Set-TransportAgent -Identity $AgentName -Priority $Priority
    
    # Enable the agent
    Enable-TransportAgent -Identity $AgentName
    
    Write-Host "Transport Agent installed successfully!" -ForegroundColor Green
    Write-Host "Agent Status:" -ForegroundColor Yellow
    Get-TransportAgent -Identity $AgentName | Format-Table Name, Enabled, Priority
    
    Write-Host "Restarting Microsoft Exchange Transport service..." -ForegroundColor Yellow
    Restart-Service MSExchangeTransport
    
    Write-Host "Installation completed successfully!" -ForegroundColor Green
    Write-Host "Monitor the agent logs at: C:\ExchangeLogs\UrlToTextAgent.log" -ForegroundColor Cyan
    
} catch {
    Write-Host "Error installing transport agent: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}