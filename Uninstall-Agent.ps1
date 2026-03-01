# Exchange Transport Agent Uninstallation Script
# Run this script as Administrator on the Exchange Server

param(
    [Parameter(Mandatory=$false)]
    [string]$AgentName = "UrlToTextAgent"
)

# Check if Exchange Management Shell is loaded
if (!(Get-Command Get-TransportAgent -ErrorAction SilentlyContinue)) {
    Write-Host "Loading Exchange Management Shell..." -ForegroundColor Yellow
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}

try {
    Write-Host "Uninstalling Transport Agent: $AgentName" -ForegroundColor Yellow
    
    # Check if agent exists
    $agent = Get-TransportAgent -Identity $AgentName -ErrorAction SilentlyContinue
    if (!$agent) {
        Write-Host "Transport Agent '$AgentName' not found." -ForegroundColor Red
        exit 1
    }
    
    # Disable the agent
    Write-Host "Disabling agent..." -ForegroundColor Yellow
    Disable-TransportAgent -Identity $AgentName -Confirm:$false
    
    # Uninstall the agent
    Write-Host "Uninstalling agent..." -ForegroundColor Yellow
    Uninstall-TransportAgent -Identity $AgentName -Confirm:$false
    
    Write-Host "Transport Agent uninstalled successfully!" -ForegroundColor Green
    
    Write-Host "Restarting Microsoft Exchange Transport service..." -ForegroundColor Yellow
    Restart-Service MSExchangeTransport
    
    Write-Host "Uninstallation completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Host "Error uninstalling transport agent: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}