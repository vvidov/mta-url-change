# URL to Text Converter Transport Agent for Exchange 2019

[![Build and Test](https://github.com/your-username/mta-url-change/actions/workflows/ci-cd.yml/badge.svg)](https://github.com/your-username/mta-url-change/actions/workflows/ci-cd.yml)
[![PowerShell Tests](https://github.com/your-username/mta-url-change/actions/workflows/powershell-test.yml/badge.svg)](https://github.com/your-username/mta-url-change/actions/workflows/powershell-test.yml)
[![Unit Tests](https://github.com/your-username/mta-url-change/actions/workflows/test.yml/badge.svg)](https://github.com/your-username/mta-url-change/actions/workflows/test.yml)

This project contains a custom Exchange Transport Agent that scans all external emails and converts clickable URLs to plain text, preserving the URL content while removing hyperlinking functionality to enhance security.

## Features

- Scans all incoming external emails
- **Automatically skips signed and encrypted emails** to preserve integrity
- Converts clickable URLs to plain text in both text and HTML email bodies
- **Multi-level logging** (DEBUG, INFO, WARNING, ERROR, SUCCESS)
- **Windows Event Log integration** for critical errors
- Configurable internal domain list
- **Process ID tracking** for debugging
- **Comprehensive error handling**
- Easy installation and uninstallation scripts

## Prerequisites

- Exchange Server 2019 Subscription Edition
- .NET Framework 4.7.2 or higher
- Visual Studio 2017/2019 or MSBuild tools
- Exchange Management Shell
- Administrator privileges on Exchange Server

## Project Structure

```
mta-url-change/
├── UrlToTextTransportAgent.csproj      # Production project (requires Exchange)
├── UrlToTextTransportAgent.Mock.csproj # Mock project (testing without Exchange)
├── UrlToTextTransportAgent.Tests.csproj # Unit test project
├── UrlToTextTransportAgent.sln         # Visual Studio solution
├── UrlToTextAgent.cs                   # Main transport agent implementation
├── UrlToTextAgentFactory.cs            # Agent factory class
├── Mock/
│   └── MockExchangeTypes.cs            # Mock Exchange API for testing
├── Tests/
│   ├── UrlToTextAgentTests.cs          # Unit tests
│   ├── TestHelpers.cs                  # Test utility functions
│   └── Properties/AssemblyInfo.cs      # Test assembly info
├── Properties/
│   ├── AssemblyInfo.cs                 # Production assembly info
│   └── AssemblyInfo.Mock.cs            # Mock assembly info
├── .github/workflows/                  # GitHub Actions CI/CD
│   ├── ci-cd.yml                       # Main build and test workflow
│   ├── test.yml                        # Unit test execution
│   ├── powershell-test.yml             # PowerShell script validation
│   └── release.yml                     # Release package creation
├── Build-Agent.ps1                     # Production build script
├── Build-Mock.ps1                      # Mock build script (for testing)
├── Install-Agent.ps1                   # Installation script
├── Uninstall-Agent.ps1                 # Uninstallation script
├── Deploy-Agent.ps1                    # Complete build and deployment
├── Test-Agent.ps1                      # Installation verification
└── README.md                           # This file
```

## Building the Agent

### Method 1: Mock Version (Development/Testing)

For development and testing without Exchange dependencies:

```powershell
# Build mock version with tests
.\Build-Mock.ps1 -RunTests

# Or just build without tests
.\Build-Mock.ps1 -Configuration Release
```

The mock version includes:
- All agent functionality with simulated Exchange API
- Full unit test coverage
- No Exchange Server dependencies
- Perfect for CI/CD and development

### Method 2: Production Build Script

For building the production version on an Exchange server:

```powershell
.\Build-Agent.ps1 -Configuration Release
```

### Method 3: Using Visual Studio

1. Open `UrlToTextTransportAgent.sln` in Visual Studio
2. Choose your target project:
   - **UrlToTextTransportAgent** - Production (requires Exchange DLLs)
   - **UrlToTextTransportAgent.Mock** - Testing (no dependencies)
   - **UrlToTextTransportAgent.Tests** - Unit tests
3. Build the solution

### Method 4: Using MSBuild directly

```powershell
# Production version (requires Exchange)
MSBuild.exe UrlToTextTransportAgent.csproj /p:Configuration=Release

# Mock version (development)  
MSBuild.exe UrlToTextTransportAgent.Mock.csproj /p:Configuration=Release

# Run tests
MSBuild.exe UrlToTextTransportAgent.Tests.csproj /p:Configuration=Release
```

### Testing

Run comprehensive unit tests:
```powershell
# Build and run all tests
.\Build-Mock.ps1 -RunTests -Configuration Debug

# Or run tests manually after building
vstest.console.exe bin\Debug\UrlToTextTransportAgent.Tests.dll
```

Test coverage includes:
- URL conversion from plain text and HTML
- Signed/encrypted email detection and skipping
- Internal vs external email classification  
- Error handling and logging
- Performance benchmarks

## CI/CD and Automation

### GitHub Actions Workflows

This project includes comprehensive GitHub Actions for automated building and testing:

**🔄 Continuous Integration** ([ci-cd.yml](.github/workflows/ci-cd.yml))
- Builds both mock and production versions
- Runs unit tests on multiple configurations
- Performs code analysis and security scanning
- Automatic release creation on version changes

**🧪 Unit Testing** ([test.yml](.github/workflows/test.yml))  
- Comprehensive unit test execution
- Performance benchmarking
- Test result reporting and artifact upload
- Daily scheduled test runs

**📜 PowerShell Validation** ([powershell-test.yml](.github/workflows/powershell-test.yml))
- Syntax validation for all PowerShell scripts
- Parameter validation testing
- Mock installation dry-run testing
- Configuration file validation

**📦 Release Packaging** ([release.yml](.github/workflows/release.yml))
- Automated release package creation
- Version management and tagging
- Multi-artifact release uploads
- Package content validation

### Local Development Workflow

```powershell
# 1. Build and test locally
.\Build-Mock.ps1 -RunTests -Configuration Debug

# 2. Test specific functionality
.\Test-Agent.ps1 -TestEmailTo "your-email@domain.com"

# 3. Deploy to test environment
.\Deploy-Agent.ps1 -AgentDllPath "bin\Release\UrlToTextTransportAgent.Mock.dll"

# 4. Build production version on Exchange server
.\Build-Agent.ps1 -Configuration Release
.\Deploy-Agent.ps1 -TestAfterInstall
```

### Automated Testing Features

- **Unit Tests**: 15+ test cases covering all functionality
- **Integration Tests**: Mock Exchange environment testing
- **Performance Tests**: Baseline performance monitoring
- **Security Tests**: Static analysis for potential secrets
- **Compatibility Tests**: Multi-configuration matrix testing

## Configuration

Before building, update the internal domains in `UrlToTextAgent.cs`:

```csharp
// Add your organization's domains here
string[] internalDomains = { "yourdomain.com", "internal.local" };
```

Replace `"yourdomain.com"` and `"internal.local"` with your actual internal domain names.

## Security Features

### Signed and Encrypted Email Protection

The agent automatically detects and **skips processing** of:

- **S/MIME signed messages** (application/pkcs7-mime, multipart/signed)
- **S/MIME encrypted messages** (enveloped-data)
- **PGP signed/encrypted messages** (-----BEGIN PGP MESSAGE-----)
- **Messages with signature attachments** (.p7s, .sig files)

This ensures that:
- Digital signatures remain valid and unbroken
- Encrypted content is not tampered with
- Email authentication mechanisms continue to work
- Compliance requirements are maintained

### Enhanced Logging

- **Multi-level logging**: DEBUG, INFO, WARNING, ERROR, SUCCESS
- **Process ID tracking** for multi-process debugging
- **Windows Event Log integration** for critical errors
- **Structured log format** with timestamps
- **Automatic log rotation** when size limits are exceeded
- **Backup logging location** if primary logging fails

Log levels can be configured in `App.config`:
```xml
<add key="LogLevel" value="INFO" />
<add key="EnableEventLog" value="true" />
<add key="SkipSignedEmails" value="true" />
<add key="SkipEncryptedEmails" value="true" />
```

## Installation

### Prerequisites Check

1. Ensure Exchange Management Shell is available
2. Verify you have Administrator privileges
3. Stop any antivirus real-time protection temporarily during installation

### Installation Steps

1. Copy the built DLL to the Exchange Server (e.g., `C:\ExchangeAgents\`)
2. Run the installation script as Administrator:
   ```powershell
   .\Install-Agent.ps1 -AgentDllPath "C:\ExchangeAgents\UrlToTextTransportAgent.dll"
   ```

### Manual Installation

If you prefer manual installation:

```powershell
# Load Exchange Management Shell
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

# Install the agent
Install-TransportAgent -Name "UrlToTextAgent" -TransportAgentFactory "UrlToTextTransportAgent.UrlToTextAgentFactory" -AssemblyPath "C:\ExchangeAgents\UrlToTextTransportAgent.dll"

# Set priority (lower number = higher priority)
Set-TransportAgent -Identity "UrlToTextAgent" -Priority 1

# Enable the agent
Enable-TransportAgent -Identity "UrlToTextAgent"

# Restart Exchange Transport Service
Restart-Service MSExchangeTransport
```

## Verification

### Check Agent Status

```powershell
Get-TransportAgent -Identity "UrlToTextAgent"
```

### Check Agent Logs

Monitor the log file at `C:\ExchangeLogs\UrlToTextAgent.log` for processing information.

### Test the Agent

1. Send a test email from an external account containing URLs
2. Check the received email to verify URLs are converted to plain text
3. Review the agent logs for processing confirmation

## Uninstallation

Use the uninstall script:

```powershell
.\Uninstall-Agent.ps1
```

Or manually:

```powershell
# Disable the agent
Disable-TransportAgent -Identity "UrlToTextAgent"

# Uninstall the agent
Uninstall-TransportAgent -Identity "UrlToTextAgent"

# Restart Exchange Transport Service
Restart-Service MSExchangeTransport
```

## Troubleshooting

### Common Issues

1. **Assembly Load Errors**
   - Ensure .NET Framework 4.7.2 is installed
   - Verify Exchange DLL references are correct
   - Check that the DLL is not blocked (Right-click → Properties → Unblock)

2. **Agent Not Processing Messages**
   - Verify the agent is enabled: `Get-TransportAgent -Identity "UrlToTextAgent"`
   - Check Exchange Transport service is running
   - Review Windows Event Logs for errors

3. **Permission Issues**
   - Ensure the agent DLL has proper permissions
   - Verify the Exchange Transport service account can access the DLL
   - Check that log directory exists and is writable

### Log Locations

- Agent logs: `C:\ExchangeLogs\UrlToTextAgent.log`
- Backup logs: `%TEMP%\UrlToTextAgent_backup.log` (if main logging fails)
- Exchange Transport logs: `C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\MessageTracking\`
- Windows Event Logs: Applications and Services Logs → Microsoft → Exchange
- **Windows Event Logs (Agent)**: Windows Logs → Application (Source: UrlToTextAgent)

### Understanding Log Levels

- **DEBUG**: Detailed processing information (URL detection, domain checks)
- **INFO**: General processing information (message processing started/completed)
- **WARNING**: Potential issues (signed/encrypted emails skipped, malformed addresses)
- **ERROR**: Processing errors that don't stop the agent
- **SUCCESS**: Successful processing with URLs converted to text

### Log Analysis Examples

```powershell
# Check for processing errors
Get-Content C:\ExchangeLogs\UrlToTextAgent.log | Where-Object { $_ -match '\[ERROR\]' }

# Count URLs converted today
$today = Get-Date -Format "yyyy-MM-dd"
Get-Content C:\ExchangeLogs\UrlToTextAgent.log | Where-Object { $_ -match $today -and $_ -match 'Total URLs converted' }

# Check signed/encrypted emails skipped
Get-Content C:\ExchangeLogs\UrlToTextAgent.log | Where-Object { $_ -match 'signed or encrypted.*skipping' }
```

## Customization

### Modifying URL Detection

Edit the `ConvertUrls()` and `ConvertHtmlUrls()` methods in `UrlToTextAgent.cs` to customize URL detection patterns.

### Changing Replacement Text

Modify the URL-to-text conversion behavior to your preferred format:

```csharp
return "[LINK BLOCKED FOR SECURITY]"; // Custom replacement text
```

### Adding Whitelist Domains

To whitelist specific domains, modify the URL processing logic to check against allowed domains before conversion.

## Security Considerations

- The agent processes all external emails, which may impact performance
- **Signed and encrypted emails are automatically skipped** to maintain integrity
- Log files may contain sensitive information - ensure proper access controls
- Regularly monitor agent performance and Exchange server health
- Consider implementing additional security measures like attachment scanning
- **Event Log integration** provides additional monitoring for security teams
- The agent **preserves email authentication** (DKIM, SPF, DMARC) by skipping signed messages

## Support

For issues or questions:
- Check Exchange Server logs and Event Viewer
- Review the agent log file for detailed processing information
- Ensure all prerequisites are met
- Test with a simple configuration first

## License

This project is provided as-is for educational and internal use purposes.