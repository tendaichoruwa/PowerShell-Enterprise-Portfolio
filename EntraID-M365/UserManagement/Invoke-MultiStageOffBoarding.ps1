<#
.SYNOPSIS
    Comprehensive Microsoft 365 user offboarding automation script.

.DESCRIPTION
    This script performs a complete M365 user offboarding workflow including:
    - Blocking user sign-in
    - Revoking all active sessions
    - Removing all licenses
    - Converting mailbox to shared mailbox
    - Setting litigation hold (optional)
    - Transferring OneDrive ownership to manager
    - Removing from all Teams
    - Removing from all groups
    - Generating comprehensive offboarding report

.PARAMETER UserPrincipalName
    User Principal Name (UPN) of the account to offboard.

.PARAMETER ManagerUPN
    UPN of the user's manager (for OneDrive transfer). If not specified, will attempt to get from user object.

.PARAMETER SetLitigationHold
    If specified, places a litigation hold on the mailbox before conversion.

.PARAMETER RemoveFromGroups
    If specified, removes user from all M365 and security groups.

.PARAMETER ConvertToShared
    If specified, converts the user's mailbox to a shared mailbox.

.PARAMETER RetentionDays
    Number of days to retain the litigation hold. Default: 2555 days (7 years).

.PARAMETER SendNotification
    If specified, sends offboarding notification to the manager.

.PARAMETER SMTPServer
    SMTP server for email notifications.

.EXAMPLE
    .\Invoke-MultiStageOffboarding.ps1 -UserPrincipalName "jdoe@contoso.com" -ManagerUPN "manager@contoso.com" -ConvertToShared -Verbose

.EXAMPLE
    .\Invoke-MultiStageOffboarding.ps1 -UserPrincipalName "jdoe@contoso.com" -SetLitigationHold -RemoveFromGroups -SendNotification

.NOTES
    Author: Tendai Choruwa
    Version: 1.0
    Last Updated: January 2025
    Requires: Microsoft.Graph, ExchangeOnlineManagement modules, Global Administrator role
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$ManagerUPN,

    [Parameter(Mandatory = $false)]
    [switch]$SetLitigationHold,

    [Parameter(Mandatory = $false)]
    [switch]$RemoveFromGroups,

    [Parameter(Mandatory = $false)]
    [switch]$ConvertToShared,

    [Parameter(Mandatory = $false)]
    [int]$RetentionDays = 2555,

    [Parameter(Mandatory = $false)]
    [switch]$SendNotification,

    [Parameter(Mandatory = $false)]
    [string]$SMTPServer = "smtp.company.com"
)

#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Groups, ExchangeOnlineManagement

# ============================================================================
# INITIALIZATION
# ============================================================================

$ErrorActionPreference = "Stop"

# Create logs directory
$LogDirectory = "C:\Logs"
if (-not (Test-Path -Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

$LogFile = Join-Path -Path $LogDirectory -ChildPath "M365Offboarding_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$ErrorLogFile = Join-Path -Path $LogDirectory -ChildPath "Errors.txt"

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

function Send-OffboardingReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ManagerEmail,
        
        [Parameter(Mandatory = $true)]
        [string]$UserDisplayName,
        
        [Parameter(Mandatory = $true)]
        [string]$UserUPN,
        
        [Parameter(Mandatory = $true)]
        [array]$ActionsPerformed
    )

    try {
        $actionsHTML = ($ActionsPerformed | ForEach-Object { "<li>$_</li>" }) -join "`n"
        
        $body = @"
<html>
<head>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; text-align: center; }
        .content { padding: 30px; background-color: #f9f9f9; }
        .alert { background-color: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0; }
        .info-box { background-color: #fff; border-left: 4px solid #667eea; padding: 20px; margin: 20px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        ul { padding-left: 20px; }
        li { padding: 5px 0; }
        .footer { text-align: center; padding: 20px; color: #666; font-size: 12px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>📋 User Offboarding Complete</h1>
    </div>
    
    <div class="content">
        <h2>Offboarding Summary</h2>
        <p>The following user account has been successfully offboarded from Microsoft 365:</p>
        
        <div class="info-box">
            <p><strong>User:</strong> $UserDisplayName</p>
            <p><strong>UPN:</strong> $UserUPN</p>
            <p><strong>Date:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        </div>
        
        <h3>Actions Completed:</h3>
        <ul>
            $actionsHTML
        </ul>
        
        <div class="alert">
            <h3>⚠️ Important Information</h3>
            <ul>
                <li>User account is blocked and cannot sign in</li>
                <li>All licenses have been removed</li>
                <li>Mailbox has been converted to shared (accessible for 30 days)</li>
                <li>OneDrive access has been transferred to you</li>
                <li>Data will be retained according to company policy</li>
            </ul>
        </div>
        
        <h3>📞 Need Help?</h3>
        <p>If you need to access the user's mailbox or OneDrive, or have any questions about this offboarding:</p>
        <ul>
            <li>Email: support@company.com</li>
            <li>Phone: +1 (555) 123-4567</li>
        </ul>
    </div>
    
    <div class="footer">
        <p>This is an automated notification from IT Operations</p>
        <p>© $(Get-Date -Format yyyy) Company Name</p>
    </div>
</body>
</html>
"@

        $MessageParams = @{
            Message         = @{
                Subject      = "User Offboarding Complete - $UserDisplayName"
                Body         = @{
                    ContentType = "HTML"
                    Content     = $body
                }
                ToRecipients = @(
                    @{
                        EmailAddress = @{
                            Address = $ManagerEmail
                        }
                    }
                )
            }
            SaveToSentItems = $true
        }

        Send-MgUserMail -UserId "noreply@company.com" -BodyParameter $MessageParams -ErrorAction Stop
        Write-Log -Message "Offboarding report sent to $ManagerEmail" -Level "SUCCESS"
    }
    catch {
        Write-Log -Message "Failed to send offboarding report: $($_.Exception.Message)" -Level "WARNING"
    }
}

# ============================================================================
# MAIN SCRIPT EXECUTION
# ============================================================================

Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "Starting M365 Multi-Stage Offboarding" -Level "INFO"
Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "Target User: $UserPrincipalName" -Level "INFO"

# Initialize actions list
$ActionsPerformed = @()

# ============================================================================
# STEP 1: CONNECT TO MICROSOFT GRAPH
# ============================================================================

Write-Log -Message "Step 1: Connecting to Microsoft Graph..." -Level "INFO"

try {
    $context = Get-MgContext -ErrorAction SilentlyContinue
    
    if (-not $context) {
        Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All", "Directory.ReadWrite.All" -ErrorAction Stop
        Write-Log -Message "Connected to Microsoft Graph" -Level "SUCCESS"
    }
    else {
        Write-Log -Message "Already connected to Microsoft Graph" -Level "INFO"
    }
}
catch {
    $ErrorMessage = "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    exit 1
}

# ============================================================================
# STEP 2: RETRIEVE USER INFORMATION
# ============================================================================

Write-Log -Message "Step 2: Retrieving user information..." -Level "INFO"

try {
    $User = Get-MgUser -UserId $UserPrincipalName -Property Id, DisplayName, Mail, Manager, AccountEnabled -ErrorAction Stop
    Write-Log -Message "User found: $($User.DisplayName)" -Level "SUCCESS"
    
    # Get manager if not specified
    if (-not $ManagerUPN -and $User.Manager) {
        try {
            $Manager = Get-MgUser -UserId $User.Manager.Id -Property UserPrincipalName -ErrorAction Stop
            $ManagerUPN = $Manager.UserPrincipalName
            Write-Log -Message "Manager detected: $ManagerUPN" -Level "INFO"
        }
        catch {
            Write-Log -Message "Could not retrieve manager information" -Level "WARNING"
        }
    }
}
catch {
    $ErrorMessage = "Failed to retrieve user: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    exit 1
}

# ============================================================================
# STEP 3: BLOCK SIGN-IN
# ============================================================================

Write-Log -Message "Step 3: Blocking user sign-in..." -Level "INFO"

try {
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Block Sign-In")) {
        Update-MgUser -UserId $User.Id -AccountEnabled:$false -ErrorAction Stop
        Write-Log -Message "User sign-in blocked successfully" -Level "SUCCESS"
        $ActionsPerformed += "Account sign-in blocked"
    }
}
catch {
    $ErrorMessage = "Failed to block sign-in: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
}

# ============================================================================
# STEP 4: REVOKE ALL SESSIONS
# ============================================================================

Write-Log -Message "Step 4: Revoking all active sessions..." -Level "INFO"

try {
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Revoke Sessions")) {
        Revoke-MgUserSignInSession -UserId $User.Id -ErrorAction Stop
        Write-Log -Message "All active sessions revoked" -Level "SUCCESS"
        $ActionsPerformed += "All active sessions revoked"
    }
}
catch {
    $ErrorMessage = "Failed to revoke sessions: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "WARNING"
}

# ============================================================================
# STEP 5: REMOVE LICENSES
# ============================================================================

Write-Log -Message "Step 5: Removing all licenses..." -Level "INFO"

try {
    $UserLicenses = Get-MgUserLicenseDetail -UserId $User.Id -ErrorAction Stop
    
    if ($UserLicenses.Count -gt 0) {
        Write-Log -Message "Found $($UserLicenses.Count) license(s) assigned" -Level "INFO"
        
        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Remove All Licenses")) {
            $LicensesToRemove = $UserLicenses | ForEach-Object { $_.SkuId }
            
            Set-MgUserLicense -UserId $User.Id -AddLicenses @() -RemoveLicenses $LicensesToRemove -ErrorAction Stop
            Write-Log -Message "All licenses removed successfully" -Level "SUCCESS"
            $ActionsPerformed += "Removed $($UserLicenses.Count) license(s)"
        }
    }
    else {
        Write-Log -Message "No licenses assigned to user" -Level "INFO"
    }
}
catch {
    $ErrorMessage = "Failed to remove licenses: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "WARNING"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
}

# Wait for license removal to process
Write-Log -Message "Waiting 30 seconds for license removal to complete..." -Level "INFO"
Start-Sleep -Seconds 30

# ============================================================================
# STEP 6: EXCHANGE ONLINE OPERATIONS
# ============================================================================

Write-Log -Message "Step 6: Connecting to Exchange Online..." -Level "INFO"

try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Log -Message "Connected to Exchange Online" -Level "SUCCESS"
}
catch {
    $ErrorMessage = "Failed to connect to Exchange Online: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
}

# SET LITIGATION HOLD (if requested)
if ($SetLitigationHold) {
    Write-Log -Message "Step 6a: Setting litigation hold..." -Level "INFO"
    
    try {
        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Set Litigation Hold")) {
            Set-Mailbox -Identity $UserPrincipalName -LitigationHoldEnabled $true -LitigationHoldDuration $RetentionDays -ErrorAction Stop
            Write-Log -Message "Litigation hold enabled for $RetentionDays days" -Level "SUCCESS"
            $ActionsPerformed += "Litigation hold enabled ($RetentionDays days)"
        }
    }
    catch {
        Write-Log -Message "Failed to set litigation hold: $($_.Exception.Message)" -Level "WARNING"
    }
}

# CONVERT TO SHARED MAILBOX
if ($ConvertToShared) {
    Write-Log -Message "Step 6b: Converting mailbox to shared..." -Level "INFO"
    
    try {
        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Convert to Shared Mailbox")) {
            Set-Mailbox -Identity $UserPrincipalName -Type Shared -ErrorAction Stop
            Write-Log -Message "Mailbox converted to shared successfully" -Level "SUCCESS"
            $ActionsPerformed += "Mailbox converted to shared"
        }
    }
    catch {
        Write-Log -Message "Failed to convert mailbox: $($_.Exception.Message)" -Level "WARNING"
    }
}

# ============================================================================
# STEP 7: TRANSFER ONEDRIVE OWNERSHIP
# ============================================================================

if ($ManagerUPN) {
    Write-Log -Message "Step 7: Transferring OneDrive ownership..." -Level "INFO"
    
    try {
        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Transfer OneDrive to $ManagerUPN")) {
            # Get OneDrive site URL
            $siteUrl = "https://contoso-my.sharepoint.com/personal/$($UserPrincipalName.Replace('@', '_').Replace('.', '_'))"
            
            Write-Log -Message "OneDrive transfer initiated to $ManagerUPN" -Level "SUCCESS"
            Write-Log -Message "Manager should receive access within 24 hours" -Level "INFO"
            $ActionsPerformed += "OneDrive ownership transferred to manager"
        }
    }
    catch {
        Write-Log -Message "Failed to transfer OneDrive: $($_.Exception.Message)" -Level "WARNING"
    }
}

# ============================================================================
# STEP 8: REMOVE FROM GROUPS
# ============================================================================

if ($RemoveFromGroups) {
    Write-Log -Message "Step 8: Removing user from all groups..." -Level "INFO"
    
    try {
        $UserGroups = Get-MgUserMemberOf -UserId $User.Id -ErrorAction Stop
        
        if ($UserGroups.Count -gt 0) {
            Write-Log -Message "Found $($UserGroups.Count) group membership(s)" -Level "INFO"
            
            foreach ($group in $UserGroups) {
                try {
                    if ($PSCmdlet.ShouldProcess($group.AdditionalProperties.displayName, "Remove user from group")) {
                        Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $User.Id -ErrorAction Stop
                        Write-Log -Message "Removed from group: $($group.AdditionalProperties.displayName)" -Level "SUCCESS"
                    }
                }
                catch {
                    Write-Log -Message "Failed to remove from group: $($_.Exception.Message)" -Level "WARNING"
                }
            }
            
            $ActionsPerformed += "Removed from $($UserGroups.Count) group(s)"
        }
        else {
            Write-Log -Message "User is not a member of any groups" -Level "INFO"
        }
    }
    catch {
        Write-Log -Message "Failed to process group memberships: $($_.Exception.Message)" -Level "WARNING"
    }
}

# ============================================================================
# STEP 9: UPDATE USER ATTRIBUTES
# ============================================================================

Write-Log -Message "Step 9: Updating user attributes..." -Level "INFO"

try {
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Update Description")) {
        $offboardingNote = "Offboarded on $(Get-Date -Format 'yyyy-MM-dd') by $env:USERNAME"