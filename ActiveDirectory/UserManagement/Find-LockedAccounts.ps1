<#
.SYNOPSIS
    Identifies and unlocks locked Active Directory user accounts across all domain controllers.

.DESCRIPTION
    This script performs comprehensive locked account detection and remediation:
    - Queries all domain controllers for locked accounts
    - Identifies the originating DC and lockout source
    - Displays bad password count and lockout time
    - Optionally unlocks accounts with confirmation
    - Sends email alerts to security team
    - Generates detailed HTML report

.PARAMETER UnlockAccounts
    If specified, automatically unlocks all found locked accounts after confirmation.

.PARAMETER EmailAlert
    If specified, sends email notification to security team with locked account details.

.PARAMETER EmailTo
    Email address(es) to send alerts to (comma-separated).

.PARAMETER SMTPServer
    SMTP server address for sending email alerts. Default: smtp.company.com

.PARAMETER OutputReport
    If specified, generates an HTML report of all locked accounts.

.PARAMETER ReportPath
    Path to save the HTML report. Default: C:\Logs\LockedAccounts_Report_YYYYMMDD.html

.PARAMETER Domain
    Fully qualified domain name. If not specified, uses current domain.

.EXAMPLE
    .\Find-LockedAccounts.ps1 -Verbose

.EXAMPLE
    .\Find-LockedAccounts.ps1 -UnlockAccounts -EmailAlert -EmailTo "security@contoso.com"

.EXAMPLE
    .\Find-LockedAccounts.ps1 -OutputReport -ReportPath "C:\Reports\LockedAccounts.html"

.NOTES
    Author: Tendai Choruwa
    Version: 1.0
    Last Updated: January 2025
    Requires: Active Directory PowerShell Module, Domain Admin or Account Operator privileges
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $false)]
    [switch]$UnlockAccounts,

    [Parameter(Mandatory = $false)]
    [switch]$EmailAlert,

    [Parameter(Mandatory = $false)]
    [string[]]$EmailTo,

    [Parameter(Mandatory = $false)]
    [string]$SMTPServer = "smtp.company.com",

    [Parameter(Mandatory = $false)]
    [switch]$OutputReport,

    [Parameter(Mandatory = $false)]
    [string]$ReportPath,

    [Parameter(Mandatory = $false)]
    [string]$Domain
)

#Requires -Modules ActiveDirectory

# ============================================================================
# INITIALIZATION
# ============================================================================

$ErrorActionPreference = "Stop"

# Create logs directory
$LogDirectory = "C:\Logs"
if (-not (Test-Path -Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

$LogFile = Join-Path -Path $LogDirectory -ChildPath "LockedAccounts_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$ErrorLogFile = Join-Path -Path $LogDirectory -ChildPath "Errors.txt"

if (-not $ReportPath) {
    $ReportPath = Join-Path -Path $LogDirectory -ChildPath "LockedAccounts_Report_$(Get-Date -Format 'yyyyMMdd').html"
}

# ============================================================================
# FUNCTIONS
# ============================================================================

function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    Add-Content -Path $LogFile -Value $LogEntry
    
    switch ($Level) {
        "INFO"    { Write-Host $LogEntry -ForegroundColor Cyan }
        "SUCCESS" { Write-Host $LogEntry -ForegroundColor Green }
        "WARNING" { Write-Host $LogEntry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $LogEntry -ForegroundColor Red }
    }
}

function Get-LockoutSource {
    <#
    .SYNOPSIS
        Identifies the source of account lockouts by querying domain controller security logs.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserName,
        
        [Parameter(Mandatory = $true)]
        [string]$DomainController
    )

    try {
        # Query security event log for lockout events (Event ID 4740)
        $FilterXML = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[(EventID=4740)]]
      and
      *[EventData[Data[@Name='TargetUserName']='$UserName']]
    </Select>
  </Query>
</QueryList>
"@

        $Events = Get-WinEvent -ComputerName $DomainController -FilterXml $FilterXML -MaxEvents 5 -ErrorAction SilentlyContinue
        
        if ($Events) {
            $LatestEvent = $Events[0]
            $EventXML = [xml]$LatestEvent.ToXml()
            
            $CallerComputer = $EventXML.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetDomainName' } | Select-Object -ExpandProperty '#text'
            
            return [PSCustomObject]@{
                SourceComputer = $CallerComputer
                LockoutTime    = $LatestEvent.TimeCreated
                EventID        = $LatestEvent.Id
            }
        }
        else {
            return $null
        }
    }
    catch {
        Write-Log -Message "Failed to query lockout source on $DomainController : $($_.Exception.Message)" -Level "WARNING"
        return $null
    }
}

function New-LockedAccountReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array]$LockedAccounts,
        
        [Parameter(Mandatory = $true)]
        [int]$TotalDCs
    )

    $accountRows = ""
    
    foreach ($account in $LockedAccounts) {
        $statusColor = if ($account.Unlocked) { "#d4edda" } else { "#f8d7da" }
        $statusBadge = if ($account.Unlocked) {
            "<span style='background-color: #28a745; color: white; padding: 4px 8px; border-radius: 4px;'>✓ Unlocked</span>"
        } else {
            "<span style='background-color: #dc3545; color: white; padding: 4px 8px; border-radius: 4px;'>🔒 Locked</span>"
        }
        
        $lockoutSource = if ($account.LockoutSource) { $account.LockoutSource.SourceComputer } else { "Unknown" }
        $lockoutTime = if ($account.LockoutTime) { $account.LockoutTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
        
        $accountRows += @"
        <tr style="background-color: $statusColor;">
            <td><strong>$($account.SamAccountName)</strong></td>
            <td>$($account.Name)</td>
            <td>$($account.Email)</td>
            <td>$lockoutTime</td>
            <td>$lockoutSource</td>
            <td>$($account.BadPwdCount)</td>
            <td>$($account.LockedOutDC)</td>
            <td style="text-align: center;">$statusBadge</td>
        </tr>
"@
    }

    $alertClass = if ($LockedAccounts.Count -gt 0) { "alert-danger" } else { "alert-success" }
    $alertMessage = if ($LockedAccounts.Count -gt 0) {
        "⚠️ ALERT: $($LockedAccounts.Count) locked account(s) detected across the domain."
    } else {
        "✓ No locked accounts found. All systems normal."
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Locked Account Detection Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            color: #333;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background-color: white;
            border-radius: 10px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }
        .header h1 {
            font-size: 32px;
            margin-bottom: 10px;
        }
        .header p {
            font-size: 16px;
            opacity: 0.9;
        }
        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 30px;
            background-color: #f8f9fa;
        }
        .summary-card {
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            text-align: center;
        }
        .summary-card h3 {
            font-size: 14px;
            color: #6c757d;
            margin-bottom: 10px;
            text-transform: uppercase;
        }
        .summary-card .value {
            font-size: 36px;
            font-weight: bold;
            color: #667eea;
        }
        .alert {
            margin: 20px 30px;
            padding: 20px;
            border-radius: 8px;
            border-left: 4px solid;
        }
        .alert-danger {
            background-color: #f8d7da;
            border-color: #dc3545;
            color: #721c24;
        }
        .alert-success {
            background-color: #d4edda;
            border-color: #28a745;
            color: #155724;
        }
        .section {
            padding: 30px;
        }
        .section h2 {
            font-size: 24px;
            margin-bottom: 20px;
            color: #667eea;
            border-bottom: 2px solid #667eea;
            padding-bottom: 10px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        th {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 12px;
        }
        td {
            padding: 12px 15px;
            border-bottom: 1px solid #dee2e6;
        }
        tr:hover {
            background-color: #f8f9fa;
        }
        .footer {
            background-color: #343a40;
            color: white;
            padding: 20px;
            text-align: center;
            font-size: 14px;
        }
        .footer p {
            margin: 5px 0;
        }
        .no-data {
            text-align: center;
            padding: 40px;
            color: #6c757d;
            font-size: 18px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔒 Locked Account Detection Report</h1>
            <p>Generated on $(Get-Date -Format "MMMM dd, yyyy 'at' HH:mm:ss")</p>
        </div>
        
        <div class="summary">
            <div class="summary-card">
                <h3>Domain Controllers Checked</h3>
                <div class="value">$TotalDCs</div>
            </div>
            <div class="summary-card">
                <h3>Locked Accounts Found</h3>
                <div class="value" style="color: #dc3545;">$($LockedAccounts.Count)</div>
            </div>
            <div class="summary-card">
                <h3>Accounts Unlocked</h3>
                <div class="value" style="color: #28a745;">$(($LockedAccounts | Where-Object { $_.Unlocked }).Count)</div>
            </div>
            <div class="summary-card">
                <h3>Scan Time</h3>
                <div class="value" style="font-size: 20px;">$(Get-Date -Format "HH:mm")</div>
            </div>
        </div>
        
        <div class="alert $alertClass">
            <strong>$alertMessage</strong>
        </div>
        
        <div class="section">
            <h2>📊 Locked Account Details</h2>
            $(if ($LockedAccounts.Count -gt 0) {
                @"
            <table>
                <thead>
                    <tr>
                        <th>Username</th>
                        <th>Display Name</th>
                        <th>Email</th>
                        <th>Lockout Time</th>
                        <th>Lockout Source</th>
                        <th>Bad Password Count</th>
                        <th>Originating DC</th>
                        <th style="text-align: center;">Status</th>
                    </tr>
                </thead>
                <tbody>
                    $accountRows
                </tbody>
            </table>
"@
            } else {
                "<div class='no-data'>✓ No locked accounts detected. All user accounts are accessible.</div>"
            })
        </div>
        
        <div class="footer">
            <p><strong>AD Account Security Monitor</strong></p>
            <p>Report generated by PowerShell automation script</p>
            <p>Log file: $LogFile</p>
            <p>© $(Get-Date -Format yyyy) IT Operations</p>
        </div>
    </div>
</body>
</html>
"@

    return $html
}

function Send-SecurityAlert {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array]$LockedAccounts,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Recipients
    )

    try {
        $accountList = ($LockedAccounts | ForEach-Object {
            "<li><strong>$($_.SamAccountName)</strong> ($($_.Name)) - Locked at: $($_.LockoutTime)</li>"
        }) -join "`n"

        $body = @"
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
        .header { background-color: #dc3545; color: white; padding: 20px; border-radius: 5px; }
        .content { padding: 20px; }
        .alert { background-color: #f8d7da; border-left: 4px solid #dc3545; padding: 15px; margin: 20px 0; }
        ul { padding-left: 20px; }
        li { padding: 5px 0; }
    </style>
</head>
<body>
    <div class="header">
        <h2>🔒 Security Alert: Locked Account Detection</h2>
    </div>
    
    <div class="content">
        <div class="alert">
            <strong>⚠️ ALERT:</strong> $($LockedAccounts.Count) user account(s) have been locked out in Active Directory.
        </div>
        
        <h3>Locked Accounts:</h3>
        <ul>
            $accountList
        </ul>
        
        <h3>Recommended Actions:</h3>
        <ul>
            <li>Verify lockout sources for potential security threats</li>
            <li>Check for password spray attacks or brute force attempts</li>
            <li>Contact users to verify legitimate lockouts</li>
            <li>Review security logs on originating domain controllers</li>
        </ul>
        
        <p><strong>Scan Time:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
        <p><strong>Full Report:</strong> $ReportPath</p>
    </div>
</body>
</html>
"@

        $EmailParams = @{
            To         = $Recipients
            From       = "security-alerts@company.com"
            Subject    = "🔒 SECURITY ALERT: $($LockedAccounts.Count) Locked Account(s) Detected"
            Body       = $body
            BodyAsHtml = $true
            SmtpServer = $SMTPServer
            Priority   = "High"
        }

        Send-MailMessage @EmailParams -ErrorAction Stop
        Write-Log -Message "Security alert email sent to: $($Recipients -join ', ')" -Level "SUCCESS"
    }
    catch {
        Write-Log -Message "Failed to send security alert email: $($_.Exception.Message)" -Level "WARNING"
    }
}

# ============================================================================
# MAIN SCRIPT EXECUTION
# ============================================================================

Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "Starting Locked Account Detection" -Level "INFO"
Write-Log -Message "========================================" -Level "INFO"

# Import Active Directory module
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log -Message "Active Directory module loaded successfully" -Level "SUCCESS"
}
catch {
    $ErrorMessage = "Failed to load Active Directory module: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    exit 1
}

# Get domain information
if (-not $Domain) {
    try {
        $Domain = (Get-ADDomain).DNSRoot
        Write-Log -Message "Using current domain: $Domain" -Level "INFO"
    }
    catch {
        Write-Log -Message "Failed to get current domain: $($_.Exception.Message)" -Level "ERROR"
        exit 1
    }
}

# ============================================================================
# STEP 1: DISCOVER ALL DOMAIN CONTROLLERS
# ============================================================================

Write-Log -Message "Step 1: Discovering domain controllers..." -Level "INFO"

try {
    $DomainControllers = Get-ADDomainController -Filter * -Server $Domain | Select-Object HostName, Site
    Write-Log -Message "Found $($DomainControllers.Count) domain controller(s)" -Level "SUCCESS"
    
    foreach ($dc in $DomainControllers) {
        Write-Log -Message "  - $($dc.HostName) (Site: $($dc.Site))" -Level "INFO"
    }
}
catch {
    $ErrorMessage = "Failed to discover domain controllers: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    exit 1
}

# ============================================================================
# STEP 2: SEARCH FOR LOCKED ACCOUNTS
# ============================================================================

Write-Log -Message "Step 2: Searching for locked accounts across all DCs..." -Level "INFO"

$LockedAccounts = @()

foreach ($dc in $DomainControllers) {
    Write-Log -Message "Querying $($dc.HostName)..." -Level "INFO"
    
    try {
        $LockedUsers = Search-ADAccount -LockedOut -Server $dc.HostName -ErrorAction Stop |
            Get-ADUser -Properties DisplayName, EmailAddress, LockedOut, LockoutTime, BadPwdCount, LastBadPasswordAttempt -ErrorAction Stop
        
        if ($LockedUsers) {
            Write-Log -Message "Found $($LockedUsers.Count) locked account(s) on $($dc.HostName)" -Level "WARNING"
            
            foreach ($user in $LockedUsers) {
                # Check if we already have this user from another DC
                if ($LockedAccounts.SamAccountName -notcontains $user.SamAccountName) {
                    
                    # Get lockout source
                    $LockoutSource = Get-LockoutSource -UserName $user.SamAccountName -DomainController $dc.HostName
                    
                    $accountInfo = [PSCustomObject]@{
                        SamAccountName = $user.SamAccountName
                        Name           = $user.DisplayName
                        Email          = $user.EmailAddress
                        LockoutTime    = $user.LockoutTime
                        BadPwdCount    = $user.BadPwdCount
                        LastBadPwdAttempt = $user.LastBadPasswordAttempt
                        LockedOutDC    = $dc.HostName
                        LockoutSource  = $LockoutSource
                        Unlocked       = $false
                    }
                    
                    $LockedAccounts += $accountInfo
                    
                    Write-Log -Message "  🔒 $($user.SamAccountName) - Locked since $($user.LockoutTime)" -Level "WARNING"
                    
                    if ($LockoutSource) {
                        Write-Log -Message "     Source: $($LockoutSource.SourceComputer)" -Level "INFO"
                    }
                }
            }
        }
        else {
            Write-Log -Message "No locked accounts found on $($dc.HostName)" -Level "SUCCESS"
        }
    }
    catch {
        $ErrorMessage = "Failed to query $($dc.HostName): $($_.Exception.Message)"
        Write-Log -Message $ErrorMessage -Level "WARNING"
        Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    }
}

Write-Log -Message "Total unique locked accounts found: $($LockedAccounts.Count)" -Level $(if ($LockedAccounts.Count -gt 0) { "WARNING" } else { "SUCCESS" })

# ============================================================================
# STEP 3: UNLOCK ACCOUNTS (IF REQUESTED)
# ============================================================================

if ($UnlockAccounts -and $LockedAccounts.Count -gt 0) {
    Write-Log -Message "Step 3: Unlocking accounts..." -Level "INFO"
    
    foreach ($account in $LockedAccounts) {
        try {
            if ($PSCmdlet.ShouldProcess($account.SamAccountName, "Unlock Account")) {
                Unlock-ADAccount -Identity $account.SamAccountName -ErrorAction Stop
                $account.Unlocked = $true
                Write-Log -Message "✓ Unlocked account: $($account.SamAccountName)" -Level "SUCCESS"
            }
        }
        catch {
            $ErrorMessage = "Failed to unlock $($account.SamAccountName): $($_.Exception.Message)"
            Write-Log -Message $ErrorMessage -Level "ERROR"
            Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
        }
    }
}

# ============================================================================
# STEP 4: GENERATE HTML REPORT
# ============================================================================

if ($OutputReport -or $LockedAccounts.Count -gt 0) {
    Write-Log -Message "Step 4: Generating HTML report..." -Level "INFO"
    
    try {
        $htmlReport = New-LockedAccountReport -LockedAccounts $LockedAccounts -TotalDCs $DomainControllers.Count
        $htmlReport | Out-File -FilePath $ReportPath -Encoding UTF8 -Force
        Write-Log -Message "HTML report saved to: $ReportPath" -Level "SUCCESS"
    }
    catch {
        $ErrorMessage = "Failed to generate HTML report: $($_.Exception.Message)"
        Write-Log -Message $ErrorMessage -Level "WARNING"
        Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    }
}

# ============================================================================
# STEP 5: SEND EMAIL ALERT (IF REQUESTED AND ACCOUNTS FOUND)
# ============================================================================

if ($EmailAlert -and $EmailTo -and $LockedAccounts.Count -gt 0) {
    Write-Log -Message "Step 5: Sending security alert email..." -Level "INFO"
    Send-SecurityAlert -LockedAccounts $LockedAccounts -Recipients $EmailTo
}

# ============================================================================
# SUMMARY REPORT
# ============================================================================

Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "Locked Account Detection Completed" -Level "INFO"
Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "Domain Controllers Checked: $($DomainControllers.Count)" -Level "INFO"
Write-Log -Message "Locked Accounts Found: $($LockedAccounts.Count)" -Level $(if ($LockedAccounts.Count -gt 0) { "WARNING" } else { "SUCCESS" })

if ($UnlockAccounts) {
    $UnlockedCount = ($LockedAccounts | Where-Object { $_.Unlocked }).Count
    Write-Log -Message "Accounts Unlocked: $UnlockedCount" -Level "SUCCESS"
}

if ($LockedAccounts.Count -gt 0) {
    Write-Log -Message "" -Level "INFO"
    Write-Log -Message "Locked Account Details:" -Level "INFO"
    
    foreach ($account in $LockedAccounts) {
        Write-Log -Message "  - $($account.SamAccountName): Locked at $($account.LockoutTime)" -Level "WARNING"
        if ($account.LockoutSource) {
            Write-Log -Message "    Source: $($account.LockoutSource.SourceComputer)" -Level "INFO"
        }
        if ($account.Unlocked) {
            Write-Log -Message "    Status: ✓ UNLOCKED" -Level "SUCCESS"
        }
    }
}

Write-Log -Message "" -Level "INFO"
Write-Log -Message "Log File: $LogFile" -Level "INFO"

if ($OutputReport -or $LockedAccounts.Count -gt 0) {
    Write-Log -Message "Report: $ReportPath" -Level "INFO"
}

# Return summary object
[PSCustomObject]@{
    ScanDate            = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Domain              = $Domain
    DomainControllersChecked = $DomainControllers.Count
    LockedAccountsFound = $LockedAccounts.Count
    AccountsUnlocked    = ($LockedAccounts | Where-Object { $_.Unlocked }).Count
    LockedAccounts      = $LockedAccounts
    ReportPath          = $ReportPath
    LogFile             = $LogFile
}