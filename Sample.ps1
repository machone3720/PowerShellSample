function Set-SessionProxy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [uri] $ProxyUri,

        [Parameter()]
        [System.Management.Automation.PSCredential] $ProxyCredential
    )
<#
.SYNOPSIS
Configures the current PowerShell session to use an HTTP/HTTPS proxy.

.DESCRIPTION
Sets .NET default web proxy, environment variables (HTTP(S)_PROXY), and
default parameters for Invoke-WebRequest/Invoke-RestMethod. Supports optional
credentialed proxies.

.PARAMETER ProxyUri
Proxy endpoint, e.g. http://proxy.contoso.com:8080

.PARAMETER ProxyCredential
Optional PSCredential for the proxy.

.EXAMPLE
Set-SessionProxy -ProxyUri http://proxy:8080

.EXAMPLE
$pc = Get-Credential
Set-SessionProxy -ProxyUri http://proxy:8080 -ProxyCredential $pc
#>
    $webProxy = New-Object System.Net.WebProxy($ProxyUri, $true)
    if ($ProxyCredential) {
        $netCred = New-Object System.Net.NetworkCredential(
            $ProxyCredential.UserName,
            $ProxyCredential.Password
        )
        $webProxy.Credentials = $netCred
    }

    # Apply to .NET defaults
    [System.Net.WebRequest]::DefaultWebProxy = $webProxy
    [System.Net.WebRequest]::DefaultWebProxy.Credentials = $webProxy.Credentials

    # Environment for tools that honor these
    $env:HTTP_PROXY  = $ProxyUri.AbsoluteUri
    $env:HTTPS_PROXY = $ProxyUri.AbsoluteUri

    # Convenience defaults for web cmdlets
    $PSDefaultParameterValues['Invoke-WebRequest:Proxy']       = $ProxyUri.AbsoluteUri
    $PSDefaultParameterValues['Invoke-RestMethod:Proxy']       = $ProxyUri.AbsoluteUri

    if ($ProxyCredential) {
        $PSDefaultParameterValues['Invoke-WebRequest:ProxyCredential'] = $ProxyCredential
        $PSDefaultParameterValues['Invoke-RestMethod:ProxyCredential'] = $ProxyCredential
    }

    Write-Verbose ("Proxy configured: {0}" -f $ProxyUri)
}

function Get-AppCredential {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Corp','Corp2','Corp3')]
        [string] $Lane,

        [Parameter()]
        [switch] $Secondary,

        [Parameter()]
        [int] $RetryCount = 3,

        [Parameter()]
        [int] $RetryDelaySeconds = 2
    )
<#
.SYNOPSIS
Builds a PSCredential from your obfuscated secrets store (Get-BankDeets).

.DESCRIPTION
Wraps your existing Get-BankDeets workflow. It reads the appropriate UPN
(PrimaryUPN/SecondaryUPN) from $Corp/$Corp2/$Corp3 and converts the returned
password into a SecureString for a PSCredential. Includes retry logic for
network file access resilience.

.PARAMETER Lane
Which lane to use (Corp, Corp2, Corp3).

.PARAMETER Secondary
Use SecondaryUPN instead of PrimaryUPN.

.PARAMETER RetryCount
Number of retry attempts for Get-BankDeets (default: 3).

.PARAMETER RetryDelaySeconds
Delay between retry attempts in seconds (default: 2).

.EXAMPLE
$cred = Get-AppCredential -Lane Corp

.EXAMPLE
$cred = Get-AppCredential -Lane Corp2 -Secondary

.EXAMPLE
$cred = Get-AppCredential -Lane Corp -RetryCount 5 -RetryDelaySeconds 3
#>
    # Populate $Corp/$Corp2/$Corp3 objects with retry logic
    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            $null = Get-BankDeets
            break
        }
        catch {
            if ($i -eq $RetryCount) {
                throw "Failed to populate lane objects after $RetryCount attempts: $($_.Exception.Message)"
            }
            Write-Verbose "Retry $i/$RetryCount: Failed to populate lane objects, retrying in $RetryDelaySeconds seconds..."
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }

    $laneObj = Get-Variable -Name $Lane -ErrorAction Stop | Select-Object -ExpandProperty Value
    $upn = if ($Secondary.IsPresent) { $laneObj.SecondaryUPN } else { $laneObj.PrimaryUPN }

    # Get plaintext with retry logic
    $plain = $null
    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            $plain = Get-BankDeets -Lane $Lane
            if ($plain) { break }
        }
        catch {
            if ($i -eq $RetryCount) {
                throw "Failed to retrieve credentials after $RetryCount attempts: $($_.Exception.Message)"
            }
            Write-Verbose "Retry $i/$RetryCount: Failed to retrieve credentials for lane '$Lane', retrying in $RetryDelaySeconds seconds..."
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }

    if (-not $plain) { throw "No password returned from Get-BankDeets for lane '$Lane' after $RetryCount attempts." }
    $secure = ConvertTo-SecureString -String $plain -AsPlainText -Force

    return New-Object System.Management.Automation.PSCredential($upn, $secure)
}

function Initialize-LogsTable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $SqlConnectionString
    )
<#
.SYNOPSIS
Ensures dbo.Logs table exists.

.DESCRIPTION
Creates dbo.Logs (if missing) with basic columns for structured script logging.

.PARAMETER SqlConnectionString
ADO.NET connection string to SQL Server.
#>
    $createSql = @"
IF NOT EXISTS (SELECT 1 FROM sys.tables t JOIN sys.schemas s ON t.schema_id=s.schema_id WHERE t.name = 'Logs' AND s.name='dbo')
BEGIN
    CREATE TABLE dbo.Logs(
        LogId            BIGINT IDENTITY(1,1) PRIMARY KEY,
        TimeUtc          DATETIME2(3) NOT NULL,
        Level            NVARCHAR(16) NOT NULL,
        Source           NVARCHAR(128) NOT NULL,
        Message          NVARCHAR(4000) NOT NULL,
        DataJson         NVARCHAR(MAX) NULL
    );
END
"@

    $conn = New-Object System.Data.SqlClient.SqlConnection($SqlConnectionString)
    $cmd  = $conn.CreateCommand()
    $cmd.CommandText = $createSql
    try {
        $conn.Open()
        [void]$cmd.ExecuteNonQuery()
    }
    finally {
        $conn.Close()
        $conn.Dispose()
    }
}

function Write-LogSql {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $SqlConnectionString,

        [Parameter(Mandatory)]
        [ValidateSet('DEBUG','INFO','WARN','ERROR')]
        [string] $Level,

        [Parameter(Mandatory)]
        [string] $Source,

        [Parameter(Mandatory)]
        [string] $Message,

        [Parameter()]
        [hashtable] $Data
    )
<#
.SYNOPSIS
Writes a structured log row to dbo.Logs.

.PARAMETER SqlConnectionString
ADO.NET connection string.

.PARAMETER Level
Severity (DEBUG, INFO, WARN, ERROR).

.PARAMETER Source
Logical source (function/module name).

.PARAMETER Message
Human-readable message.

.PARAMETER Data
Optional hashtable of extra fields; serialized to JSON.
#>
    $json = if ($Data) { $Data | ConvertTo-Json -Depth 6 -Compress } else { $null }

    $insertSql = @"
INSERT INTO dbo.Logs (TimeUtc, Level, Source, Message, DataJson)
VALUES (@TimeUtc, @Level, @Source, @Message, @DataJson);
"@

    $conn = New-Object System.Data.SqlClient.SqlConnection($SqlConnectionString)
    $cmd  = $conn.CreateCommand()
    $cmd.CommandText = $insertSql

    $null = $cmd.Parameters.Add("@TimeUtc",   [System.Data.SqlDbType]::DateTime2)
    $null = $cmd.Parameters.Add("@Level",     [System.Data.SqlDbType]::NVarChar, 16)
    $null = $cmd.Parameters.Add("@Source",    [System.Data.SqlDbType]::NVarChar, 128)
    $null = $cmd.Parameters.Add("@Message",   [System.Data.SqlDbType]::NVarChar, 4000)
    $null = $cmd.Parameters.Add("@DataJson",  [System.Data.SqlDbType]::NVarChar, -1)

    $cmd.Parameters["@TimeUtc"].Value  = [DateTime]::UtcNow
    $cmd.Parameters["@Level"].Value    = $Level
    $cmd.Parameters["@Source"].Value   = $Source
    $cmd.Parameters["@Message"].Value  = $Message
    $cmd.Parameters["@DataJson"].Value = if ($json) { $json } else { [DBNull]::Value }

    try {
        $conn.Open()
        [void]$cmd.ExecuteNonQuery()
    }
    finally {
        $conn.Close()
        $conn.Dispose()
    }
}

function Connect-ExchangeOnlineUnattended {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string] $TenantId,             # <tenant>.onmicrosoft.com or GUID

        [Parameter()]
        [string] $AppId,                # App registration (preferred for unattended)

        [Parameter()]
        [string] $CertificateThumbprint,

        [Parameter()]
        [System.Management.Automation.PSCredential] $FallbackUserCredential,  # optional: your obfuscated creds

        [Parameter()]
        [uri] $ProxyUri,

        [Parameter()]
        [System.Management.Automation.PSCredential] $ProxyCredential
    )
<#
.SYNOPSIS
Connects to Exchange Online with app-based auth if available; otherwise falls back to user creds.
Honors session proxy settings if provided.

.DESCRIPTION
Prefers certificate-based app auth for unattended runs. If AppId/Thumbprint/TenantId are not provided,
will fall back to Connect-ExchangeOnline with PSCredential (e.g., from Get-AppCredential).
Optionally configures a session proxy first.

.PARAMETER TenantId
Azure AD tenant (domain or GUID).

.PARAMETER AppId
Client ID of app registration.

.PARAMETER CertificateThumbprint
Thumbprint of cert installed in CurrentUser\My or LocalMachine\My.

.PARAMETER FallbackUserCredential
Optional PSCredential for user-based auth.

.PARAMETER ProxyUri
Optional proxy to set for the session.

.PARAMETER ProxyCredential
Optional proxy credential.
#>
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    if ($ProxyUri) {
        Set-SessionProxy -ProxyUri $ProxyUri -ProxyCredential $ProxyCredential
    }

    $info = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($info -and $info.State -eq 'Connected') { return }

    if ($TenantId -and $AppId -and $CertificateThumbprint) {
        Connect-ExchangeOnline `
            -AppId $AppId `
            -Organization $TenantId `
            -CertificateThumbprint $CertificateThumbprint `
            -ShowBanner:$false
    }
    elseif ($FallbackUserCredential) {
        Connect-ExchangeOnline `
            -Credential $FallbackUserCredential `
            -ShowBanner:$false
    }
    else {
        throw "No valid auth parameters provided. Supply (TenantId+AppId+CertificateThumbprint) or FallbackUserCredential."
    }
}

function Set-RemoteMailboxBaseline {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Low')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, Position=0)]
        [ValidateNotNullOrEmpty()]
        [string] $Identity,  # sAMAccountName or alias (on-prem)

        [Parameter(Mandatory)]
        [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
        [string] $PrimarySmtpAddress,

        [Parameter(Mandatory)]
        [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
        [string] $RemoteRoutingAddress,  # e.g. user@contoso.mail.onmicrosoft.com

        [Parameter()]
        [switch] $EnableEmailAddressPolicy,

        [Parameter()]
        [string] $RetentionPolicy = 'Corp-Default',

        [Parameter()]
        [string] $SqlConnectionString,   # if provided, function will log to dbo.Logs

        [Parameter()]
        [string] $LogSource = 'Set-RemoteMailboxBaseline'
    )
<#
.SYNOPSIS
Applies a standardized baseline to a hybrid remote mailbox (on-prem side) and logs actions to SQL.

.DESCRIPTION
Idempotently ensures a user has a Remote Mailbox and applies baseline settings:
- Ensures Remote Mailbox exists.
- Sets Primary SMTP and Remote Routing address.
- Optionally enables Email Address Policy.
- Applies a Retention Policy (if managed on-prem in your org).
Writes INFO/ERROR rows to dbo.Logs when -SqlConnectionString is provided.

.PARAMETER Identity
On-prem identity (e.g., sAMAccountName, alias, or DN).

.PARAMETER PrimarySmtpAddress
Primary email address.

.PARAMETER RemoteRoutingAddress
Target routing address (*.mail.onmicrosoft.com).

.PARAMETER EnableEmailAddressPolicy
If specified, enables Email Address Policy.

.PARAMETER RetentionPolicy
Policy to apply (default: Corp-Default).

.PARAMETER SqlConnectionString
If provided, logs to dbo.Logs (table auto-created if needed).

.PARAMETER LogSource
Source value for log rows (default: Set-RemoteMailboxBaseline).

.EXAMPLE
Set-RemoteMailboxBaseline -Identity jsmith -PrimarySmtpAddress jsmith@contoso.com -RemoteRoutingAddress jsmith@contoso.mail.onmicrosoft.com -SqlConnectionString "Server=.;Database=Ops;Integrated Security=SSPI;TrustServerCertificate=True"
#>
    try {
        if ($SqlConnectionString) { Initialize-LogsTable -SqlConnectionString $SqlConnectionString }

        Write-Verbose ("Resolving identity: {0}" -f $Identity)

        $user = Get-User -Identity $Identity -ErrorAction SilentlyContinue
        if (-not $user) {
            throw "User '$Identity' not found in on-prem directory."
        }

        $remoteMbx = Get-RemoteMailbox -Identity $Identity -ErrorAction SilentlyContinue
        if (-not $remoteMbx) {
            if ($PSCmdlet.ShouldProcess($Identity, "Enable-RemoteMailbox $PrimarySmtpAddress")) {
                Enable-RemoteMailbox -Identity $Identity -PrimarySmtpAddress $PrimarySmtpAddress -ErrorAction Stop
                $remoteMbx = Get-RemoteMailbox -Identity $Identity -ErrorAction Stop
                if ($SqlConnectionString) {
                    Write-LogSql -SqlConnectionString $SqlConnectionString -Level INFO -Source $LogSource -Message "Enabled RemoteMailbox" -Data @{ Identity = $Identity; PrimarySmtp = $PrimarySmtpAddress }
                }
            }
        }

        if ($remoteMbx.PrimarySmtpAddress.ToString() -ne $PrimarySmtpAddress) {
            if ($PSCmdlet.ShouldProcess($Identity, "Set PrimarySmtpAddress=$PrimarySmtpAddress")) {
                Set-RemoteMailbox -Identity $Identity -PrimarySmtpAddress $PrimarySmtpAddress -ErrorAction Stop
                if ($SqlConnectionString) {
                    Write-LogSql -SqlConnectionString $SqlConnectionString -Level INFO -Source $LogSource -Message "Updated PrimarySmtpAddress" -Data @{ Identity = $Identity; PrimarySmtp = $PrimarySmtpAddress }
                }
            }
        }

        if ($remoteMbx.RemoteRoutingAddress.ToString() -ne $RemoteRoutingAddress) {
            if ($PSCmdlet.ShouldProcess($Identity, "Set RemoteRoutingAddress=$RemoteRoutingAddress")) {
                Set-RemoteMailbox -Identity $Identity -RemoteRoutingAddress $RemoteRoutingAddress -ErrorAction Stop
                if ($SqlConnectionString) {
                    Write-LogSql -SqlConnectionString $SqlConnectionString -Level INFO -Source $LogSource -Message "Updated RemoteRoutingAddress" -Data @{ Identity = $Identity; RemoteRouting = $RemoteRoutingAddress }
                }
            }
        }

        if ($EnableEmailAddressPolicy.IsPresent -and -not $remoteMbx.EmailAddressPolicyEnabled) {
            if ($PSCmdlet.ShouldProcess($Identity, "Enable EmailAddressPolicy")) {
                Set-RemoteMailbox -Identity $Identity -EmailAddressPolicyEnabled:$true -ErrorAction Stop
                if ($SqlConnectionString) {
                    Write-LogSql -SqlConnectionString $SqlConnectionString -Level INFO -Source $LogSource -Message "Enabled EmailAddressPolicy" -Data @{ Identity = $Identity }
                }
            }
        }

        if ($RetentionPolicy) {
            $current = (Get-RemoteMailbox -Identity $Identity).RetentionPolicy
            if ($current -ne $RetentionPolicy) {
                if ($PSCmdlet.ShouldProcess($Identity, "Apply RetentionPolicy=$RetentionPolicy")) {
                    Set-RemoteMailbox -Identity $Identity -RetentionPolicy $RetentionPolicy -ErrorAction SilentlyContinue
                    if ($SqlConnectionString) {
                        Write-LogSql -SqlConnectionString $SqlConnectionString -Level INFO -Source $LogSource -Message "Applied RetentionPolicy" -Data @{ Identity = $Identity; RetentionPolicy = $RetentionPolicy }
                    }
                }
            }
        }

        $result = [pscustomobject]@{
            Identity              = $Identity
            PrimarySmtpAddress    = $PrimarySmtpAddress
            RemoteRoutingAddress  = $RemoteRoutingAddress
            EmailPolicyEnabled    = (Get-RemoteMailbox -Identity $Identity).EmailAddressPolicyEnabled
            RetentionPolicy       = (Get-RemoteMailbox -Identity $Identity).RetentionPolicy
            Timestamp             = Get-Date
        }

        if ($SqlConnectionString) {
            Write-LogSql -SqlConnectionString $SqlConnectionString -Level INFO -Source $LogSource -Message "Completed baseline" -Data @{ Identity = $Identity }
        }

        return $result
    }
    catch {
        if ($SqlConnectionString) {
            Write-LogSql -SqlConnectionString $SqlConnectionString -Level ERROR -Source $LogSource -Message ("Error: " + $_.Exception.Message) -Data @{ Identity = $Identity; Stack = $_.ScriptStackTrace }
        }
        throw
    }
}