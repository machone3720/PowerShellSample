# PowerShell Exchange Automation Sample

A collection of PowerShell functions designed for Exchange hybrid environment automation, showcasing enterprise-grade patterns for credential management, logging, and Exchange Online/on-premises operations.

## Overview

This sample demonstrates PowerShell best practices for Exchange administration in hybrid environments, including:

- **Secure credential management** with retry logic for network resilience
- **Certificate-based authentication** for unattended operations
- **Structured SQL logging** for audit trails and troubleshooting
- **Hybrid Exchange operations** (on-premises and Exchange Online)
- **Enterprise proxy support** for corporate network environments

## Functions

### `Set-SessionProxy`
Configures PowerShell session to use HTTP/HTTPS proxy with optional authentication.

### `Get-AppCredential`
Retrieves credentials from secure network store with configurable retry logic for resilience against transient network issues.

### `Initialize-LogsTable` / `Write-LogSql`
Creates and manages structured logging to SQL Server with JSON metadata support.

### `Connect-ExchangeOnlineUnattended`
Establishes Exchange Online connection using certificate-based app authentication or credential fallback, with proxy support.

### `Set-RemoteMailboxBaseline`
Applies standardized configuration to hybrid remote mailboxes with comprehensive logging and `-WhatIf` support.

## Key Features

- **Production-ready error handling** with meaningful error messages
- **Idempotent operations** safe for repeated execution
- **Comprehensive parameter validation** with appropriate types and patterns
- **Verbose logging** for troubleshooting and audit requirements
- **SQL-based structured logging** for enterprise monitoring
- **Certificate-based authentication** for secure unattended operations
- **Network resilience** with configurable retry logic

## Prerequisites

- PowerShell 5.1 or PowerShell 7+
- Exchange Management Shell (for on-premises operations)
- ExchangeOnlineManagement module (for cloud operations)
- SQL Server connectivity (for logging features)
- Appropriate Exchange permissions in hybrid environment

## Usage Example

```powershell
# Connect to Exchange Online with certificate auth
Connect-ExchangeOnlineUnattended -TenantId "contoso.onmicrosoft.com" -AppId "12345678-1234-1234-1234-123456789012" -CertificateThumbprint "A1B2C3D4E5F6..."

# Apply baseline configuration to remote mailbox with logging
$connString = "Server=sqlserver;Database=ExchangeOps;Integrated Security=SSPI;TrustServerCertificate=True"
Set-RemoteMailboxBaseline -Identity "jsmith" -PrimarySmtpAddress "jsmith@contoso.com" -RemoteRoutingAddress "jsmith@contoso.mail.onmicrosoft.com" -SqlConnectionString $connString -EnableEmailAddressPolicy
```

## Security Considerations

- Uses certificate-based authentication for production scenarios
- Implements secure credential storage patterns
- Includes comprehensive parameter validation
- Supports enterprise proxy configurations
- Maintains audit trails through structured logging

## Target Environment

Designed for Exchange hybrid environments with:
- Exchange Server 2019 on-premises
- Exchange Online (Microsoft 365)
- Enterprise network with proxy requirements
- SQL Server for centralized logging
- Certificate-based service authentication

---

*This sample code demonstrates PowerShell automation patterns suitable for enterprise Exchange administration roles.*