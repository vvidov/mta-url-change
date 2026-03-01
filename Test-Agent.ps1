# Test script for URL-to-Text Transport Agent
# This script helps verify that the agent is working correctly

param(
    [Parameter(Mandatory=$false)]
    [string]$AgentName = "UrlToTextAgent",
    
    [Parameter(Mandatory=$false)]
    [string]$TestEmailTo = "",
    
    [Parameter(Mandatory=$false)]
    [string]$SmtpServer = "localhost"
)

# Check if Exchange Management Shell is loaded
if (!(Get-Command Get-TransportAgent -ErrorAction SilentlyContinue)) {
    Write-Host "Loading Exchange Management Shell..." -ForegroundColor Yellow
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}

Write-Host "URL-to-Text Transport Agent Test Script" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Green

# Check agent status
Write-Host "`n1. Checking Transport Agent Status..." -ForegroundColor Yellow
try {
    $agent = Get-TransportAgent -Identity $AgentName -ErrorAction SilentlyContinue
    if ($agent) {
        Write-Host "Agent Status:" -ForegroundColor Cyan
        $agent | Format-Table Name, Enabled, Priority, TransportAgentFactory
        
        if ($agent.Enabled) {
            Write-Host "✓ Agent is enabled and ready" -ForegroundColor Green
        } else {
            Write-Host "✗ Agent is installed but disabled" -ForegroundColor Red
            Write-Host "Run: Enable-TransportAgent -Identity '$AgentName'" -ForegroundColor Yellow
        }
    } else {
        Write-Host "✗ Agent '$AgentName' not found" -ForegroundColor Red
        Write-Host "Please install the agent first using Install-Agent.ps1" -ForegroundColor Yellow
        exit 1
    }
} catch {
    Write-Host "✗ Error checking agent status: $($_.Exception.Message)" -ForegroundColor Red
}

# Check Exchange Transport Service
Write-Host "`n2. Checking Exchange Transport Service..." -ForegroundColor Yellow
try {
    $service = Get-Service MSExchangeTransport -ErrorAction SilentlyContinue
    if ($service) {
        Write-Host "Service Status: $($service.Status)" -ForegroundColor Cyan
        if ($service.Status -eq "Running") {
            Write-Host "✓ Exchange Transport Service is running" -ForegroundColor Green
        } else {
            Write-Host "✗ Exchange Transport Service is not running" -ForegroundColor Red
            Write-Host "Run: Start-Service MSExchangeTransport" -ForegroundColor Yellow
        }
    } else {
        Write-Host "✗ Exchange Transport Service not found" -ForegroundColor Red
    }
} catch {
    Write-Host "✗ Error checking service status: $($_.Exception.Message)" -ForegroundColor Red
}

# Check log file
Write-Host "`n3. Checking Agent Log File..." -ForegroundColor Yellow
# Try to get log path from environment or use default
$logPath = if ($env:EXCHANGE_AGENT_LOG_PATH) { 
    $env:EXCHANGE_AGENT_LOG_PATH 
} else { 
    "$env:ProgramData\Microsoft\Exchange\Logs\UrlToTextAgent.log" 
}
try {
    if (Test-Path $logPath) {
        $logInfo = Get-Item $logPath
        Write-Host "Log file exists: $logPath" -ForegroundColor Cyan
        Write-Host "Size: $([math]::Round($logInfo.Length / 1KB, 2)) KB" -ForegroundColor Cyan
        Write-Host "Last Modified: $($logInfo.LastWriteTime)" -ForegroundColor Cyan
        
        # Show last 10 log entries with better formatting
        Write-Host "`nLast 10 log entries:" -ForegroundColor Cyan
        Get-Content $logPath -Tail 10 | ForEach-Object { 
            if ($_ -match '\[ERROR\]') {
                Write-Host "  $_" -ForegroundColor Red
            } elseif ($_ -match '\[WARNING\]') {
                Write-Host "  $_" -ForegroundColor Yellow
            } elseif ($_ -match '\[SUCCESS\]') {
                Write-Host "  $_" -ForegroundColor Green
            } else {
                Write-Host "  $_" -ForegroundColor Gray
            }
        }
        
        Write-Host "`nLog Statistics:" -ForegroundColor Cyan
        $logContent = Get-Content $logPath
        $errorCount = ($logContent | Where-Object { $_ -match '\[ERROR\]' }).Count
        $warningCount = ($logContent | Where-Object { $_ -match '\[WARNING\]' }).Count
        $successCount = ($logContent | Where-Object { $_ -match '\[SUCCESS\]' }).Count
        $signedSkipped = ($logContent | Where-Object { $_ -match 'signed or encrypted.*skipping' }).Count
        
        Write-Host "  Total Errors: $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
        Write-Host "  Total Warnings: $warningCount" -ForegroundColor $(if ($warningCount -gt 0) { "Yellow" } else { "Green" })
        Write-Host "  Successful Processings: $successCount" -ForegroundColor Green
        Write-Host "  Signed/Encrypted Emails Skipped: $signedSkipped" -ForegroundColor Cyan
        
        Write-Host "✓ Log file is accessible" -ForegroundColor Green
    } else {
        Write-Host "⚠ Log file not found (this is normal if no messages have been processed yet)" -ForegroundColor Yellow
        Write-Host "Expected location: $logPath" -ForegroundColor Gray
    }
} catch {
    Write-Host "✗ Error accessing log file: $($_.Exception.Message)" -ForegroundColor Red
}

# Check message tracking logs
Write-Host "`n4. Checking Recent Message Flow..." -ForegroundColor Yellow
try {
    $yesterday = (Get-Date).AddDays(-1)
    $messages = Get-MessageTrackingLog -Start $yesterday -ResultSize 10 -EventId "RECEIVE" | Select-Object Timestamp, Sender, Recipients, MessageSubject
    
    if ($messages) {
        Write-Host "Recent external messages (last 10):" -ForegroundColor Cyan
        $messages | Format-Table Timestamp, Sender, @{Name="Recipients"; Expression={$_.Recipients -join ", "}}, MessageSubject
    } else {
        Write-Host "No recent messages found in tracking logs" -ForegroundColor Yellow
    }
} catch {
    Write-Host "Could not retrieve message tracking logs: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Offer to send test email
if (![string]::IsNullOrEmpty($TestEmailTo)) {
    Write-Host "`n5. Sending Test Email..." -ForegroundColor Yellow
    try {
        $testSubject = "URL-to-Text Agent Test - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $testBody = @"
This is a test email to verify the URL-to-Text Agent is working correctly.

The following URLs should be converted to plain text (non-clickable):

Plain text URLs:
- https://www.example.com
- http://malicious-site.com/phishing
- https://github.com/microsoft/exchange-server

HTML links:
<a href="https://www.microsoft.com">Microsoft Website</a>
<a href="http://suspicious-link.com">Click here for free money!</a>

If you receive this email with URLs as plain text (not clickable), the agent is working correctly.
If URLs are still clickable hyperlinks, the agent is not working properly.

NOTE: This test email should be processed since it's plain text.
Signed or encrypted emails will be skipped automatically to preserve their integrity.
"@

        Send-MailMessage -To $TestEmailTo -From "test@external-domain.com" -Subject $testSubject -Body $testBody -SmtpServer $SmtpServer -BodyAsHtml:$false
        
        Write-Host "✓ Test email sent to $TestEmailTo" -ForegroundColor Green
        Write-Host "Check the received email to verify URL conversion is working" -ForegroundColor Cyan
        
        # Also send an HTML test email
        $htmlTestSubject = "URL-to-Text Agent HTML Test - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $htmlTestBody = @"
<html>
<body>
<h3>URL-to-Text Agent HTML Test</h3>
<p>This is an HTML test email to verify the URL-to-Text Agent is working correctly.</p>

<p>The following URLs should be replaced:</p>
<ul>
<li>Plain URL: https://www.example.com</li>
<li>HTML Link: <a href="https://www.microsoft.com">Microsoft Website</a></li>
<li>Another Link: <a href="http://suspicious-link.com">Click here for free money!</a></li>
</ul>

<p><b>Note:</b> Signed or encrypted emails are automatically skipped to preserve security.</p>
</body>
</html>
"@
        
        Send-MailMessage -To $TestEmailTo -From "htmltest@external-domain.com" -Subject $htmlTestSubject -Body $htmlTestBody -SmtpServer $SmtpServer -BodyAsHtml:$true
        
        Write-Host "✓ HTML test email also sent to $TestEmailTo" -ForegroundColor Green
        
    } catch {
        Write-Host "✗ Error sending test email: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "6. Summary and Next Steps" -ForegroundColor Yellow
Write-Host "========================" -ForegroundColor Yellow

if ($agent -and $agent.Enabled -and $service -and $service.Status -eq "Running") {
    Write-Host "✓ Transport Agent appears to be configured correctly!" -ForegroundColor Green
    Write-Host "`nNew Security Features:" -ForegroundColor Cyan
    Write-Host "• Signed emails (S/MIME, PGP) are automatically skipped" -ForegroundColor Gray
    Write-Host "• Encrypted emails are automatically skipped" -ForegroundColor Gray
    Write-Host "• Enhanced logging with different severity levels" -ForegroundColor Gray
    Write-Host "• Process ID tracking for debugging" -ForegroundColor Gray
    Write-Host "• Windows Event Log integration for errors" -ForegroundColor Gray
    
    Write-Host "`nTo test the agent:" -ForegroundColor Cyan
    Write-Host "1. Send an email from an external account containing URLs" -ForegroundColor Gray
    Write-Host "2. Check if URLs are converted to plain text (no longer clickable)" -ForegroundColor Gray
    Write-Host "3. Monitor the log file: $logPath" -ForegroundColor Gray
    Write-Host "4. Send a signed email to verify it's skipped" -ForegroundColor Gray
    Write-Host "5. Use Message Tracking to verify processing" -ForegroundColor Gray
} else {
    Write-Host "⚠ Configuration issues detected. Please review the output above." -ForegroundColor Yellow
}

Write-Host "`nFor troubleshooting, check:" -ForegroundColor Cyan
Write-Host "- Windows Event Viewer (Application and System logs)" -ForegroundColor Gray
Write-Host "- Exchange Admin Center > Mail flow > Transport rules" -ForegroundColor Gray
Write-Host "- Get-TransportAgent | Format-Table Name, Enabled, Priority" -ForegroundColor Gray
Write-Host "- Test-ServiceHealth MSExchangeTransport" -ForegroundColor Gray