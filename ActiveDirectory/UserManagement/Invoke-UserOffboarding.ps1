<#
.SYNOPSIS
    Comprehensive Active Directory user offboarding automation script.

.DESCRIPTION
    This script performs a complete AD account deactivation workflow including:
    - Disabling the user account
    - Resetting password to a random value
    - Removing all group memberships (except Domain Users)
    - Moving account to Disabled Users OU
    - Setting account expiration date
    - Documenting all actions taken

.PARAMETER UserPrincipalName
    User Principal Name (UPN) of the account to offboard.

.PARAMETER DisabledOU
    Distinguished Name of the Disabled Users OU. If not specified, uses default OU.

.PARAMETER RemoveGroups
    If specified, removes user from all groups except Domain Users.

.PARAMETER SetExpiration
    If specified, sets account expiration date to 90 days from today.

.PARAMETER ExpirationDays
    Number of days until account expires. Default: 90 days.

.PARAMETER NotifyManager
    If specified, sends notification email to user's manager.

.PARAMETER SMTPServer
    SMTP server for email notifications. Default: smtp.company.com

.PARAMETER HideFromGAL
    If specified, hides the mailbox from Global Address List.

.EXAMPLE
    .\Invoke-UserOffboarding.ps1 -UserPrincipalName "jdoe@contoso.com" -DisabledOU "OU=Disabled,DC=contoso,DC=com"

.EXAMPLE
    .\Invoke-UserOffboarding.ps1 -UserPrincipalName "jdoe@contoso.com" -RemoveGroups -SetExpiration -NotifyManager -Verbose

.NOTES
    Author: Tendai Choruwa
    Version: 1.0
    Last Updated: January 2025
    Requires: Active Directory PowerShell Module, Domain Admin privileges
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$DisabledOU,

    [Parameter(Mandatory = $false)]
    [switch]$RemoveGroups,

    [Parameter(Mandatory = $false)]
    [switch]$SetExpiration,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 365)]
    [int]$ExpirationDays = 90,

    [Parameter(Mandatory = $false)]
    [switch]$NotifyManager,

    [Parameter(Mandatory = $false)]
    [string]$SMTPServer = "smtp.company.com",

    [Parameter(Mandatory = $false)]
    [switch]$HideFromGAL
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

$LogFile = Join-Path -Path $LogDirectory -ChildPath "UserOffboarding_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
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

function New-RandomPassword {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [int]$Length = 32
    )

    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_+-=[]{}|;:,.<>?"
    $password = -join ((1..$Length) | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
    return $password
}

function Send-ManagerNotification {
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
        $Subject = "User Account Offboarding Notification - $UserDisplayName"
        
        $ActionsHTML = ($ActionsPerformed | ForEach-Object { "<li>$_</li>" }) -join "`n"
        
        $Body = @"
<html>
<body style='font-family: Arial, sans-serif;'>
    <h2>User Account Offboarding Notification</h2>
    <p>This is to notify you that the following user account has been offboarded:</p>
    
    <div style='background-color: #fff3cd; padding: 15px; border-left: 4px solid #ffc107; margin: 20px 0;'>
        <p><strong>User:</strong> $UserDisplayName</p>
        <p><strong>UPN:</strong> $UserUPN</p>
        <p><strong>Date:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    </div>
    
    <h3>Actions Performed:</h3>
    <ul>
        $ActionsHTML
    </ul>
    
    <p><strong>Note:</strong> The account has been disabled and will be retained according to company policy.</p>
    
    <hr>
    <p style='font-size: 12px; color: #666;'>This is an automated notification from IT Operations. For questions, contact IT Support.</p>
</body>
</html>
"@

        $EmailParams = @{
            To         = $ManagerEmail
            From       = "itops@company.com"
            Subject    = $Subject
            Body       = $Body
            BodyAsHtml = $true
            SmtpServer = $SMTPServer
        }

        Send-MailMessage @EmailParams -ErrorAction Stop
        Write-Log -Message "Manager notification sent to $ManagerEmail" -Level "SUCCESS"
    }
    catch {
        Write-Log -Message "Failed to send manager notification: $($_.Exception.Message)" -Level "WARNING"
    }
}

# ============================================================================
# MAIN SCRIPT EXECUTION
# ============================================================================

Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "Starting User Offboarding Process" -Level "INFO"
Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "Target User: $UserPrincipalName" -Level "INFO"

# Import Active Directory module
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log -Message "Active Directory module loaded successfully" -Level "SUCCESS"
}
catch {
    Write-Log -Message "Failed to load Active Directory module: $($_.Exception.Message)" -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - Failed to load AD module: $($_.Exception.Message)"
    exit 1
}

# Retrieve user object
try {
    $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$UserPrincipalName'" -Properties * -ErrorAction Stop
    
    if (-not $ADUser) {
        throw "User not found: $UserPrincipalName"
    }
    
    Write-Log -Message "User found: $($ADUser.DisplayName) ($($ADUser.SamAccountName))" -Level "SUCCESS"
}
catch {
    $ErrorMessage = "Failed to retrieve user: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    exit 1
}

# Initialize actions list for reporting
$ActionsPerformed = @()

# Check if user is already disabled
if (-not $ADUser.Enabled) {
    Write-Log -Message "WARNING: User account is already disabled" -Level "WARNING"
}

# ============================================================================
# STEP 1: DISABLE ACCOUNT
# ============================================================================

Write-Log -Message "Step 1: Disabling user account..." -Level "INFO"

try {
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Disable Account")) {
        Disable-ADAccount -Identity $ADUser.DistinguishedName -ErrorAction Stop
        Write-Log -Message "Account disabled successfully" -Level "SUCCESS"
        $ActionsPerformed += "Account disabled"
    }
}
catch {
    $ErrorMessage = "Failed to disable account: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
}

# ============================================================================
# STEP 2: RESET PASSWORD
# ============================================================================

Write-Log -Message "Step 2: Resetting password to random value..." -Level "INFO"

try {
    $NewPassword = New-RandomPassword -Length 32
    $SecurePassword = ConvertTo-SecureString -String $NewPassword -AsPlainText -Force
    
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Reset Password")) {
        Set-ADAccountPassword -Identity $ADUser.DistinguishedName -NewPassword $SecurePassword -Reset -ErrorAction Stop
        Write-Log -Message "Password reset successfully" -Level "SUCCESS"
        $ActionsPerformed += "Password reset to random value"
    }
    
    # Clear password from memory
    $NewPassword = $null
    $SecurePassword = $null
}
catch {
    $ErrorMessage = "Failed to reset password: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
}

# ============================================================================
# STEP 3: REMOVE GROUP MEMBERSHIPS
# ============================================================================

if ($RemoveGroups) {
    Write-Log -Message "Step 3: Removing group memberships..." -Level "INFO"
    
    try {
        $UserGroups = Get-ADPrincipalGroupMembership -Identity $ADUser.DistinguishedName -ErrorAction Stop
        $GroupsToRemove = $UserGroups | Where-Object { $_.Name -ne "Domain Users" }
        
        if ($GroupsToRemove.Count -gt 0) {
            Write-Log -Message "Found $($GroupsToRemove.Count) groups to remove" -Level "INFO"
            
            foreach ($Group in $GroupsToRemove) {
                try {
                    if ($PSCmdlet.ShouldProcess("$($Group.Name)", "Remove user from group")) {
                        Remove-ADGroupMember -Identity $Group.DistinguishedName -Members $ADUser.DistinguishedName -Confirm:$false -ErrorAction Stop
                        Write-Log -Message "Removed from group: $($Group.Name)" -Level "SUCCESS"
                    }
                }
                catch {
                    Write-Log -Message "Failed to remove from group '$($Group.Name)': $($_.Exception.Message)" -Level "WARNING"
                }
            }
            
            $ActionsPerformed += "Removed from $($GroupsToRemove.Count) security groups"
        }
        else {
            Write-Log -Message "No groups to remove (only Domain Users membership)" -Level "INFO"
        }
    }
    catch {
        $ErrorMessage = "Failed to process group memberships: $($_.Exception.Message)"
        Write-Log -Message $ErrorMessage -Level "ERROR"
        Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    }
}

# ============================================================================
# STEP 4: SET ACCOUNT EXPIRATION
# ============================================================================

if ($SetExpiration) {
    Write-Log -Message "Step 4: Setting account expiration date..." -Level "INFO"
    
    try {
        $ExpirationDate = (Get-Date).AddDays($ExpirationDays)
        
        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Set expiration to $ExpirationDate")) {
            Set-ADAccountExpiration -Identity $ADUser.DistinguishedName -DateTime $ExpirationDate -ErrorAction Stop
            Write-Log -Message "Account expiration set to: $($ExpirationDate.ToString('yyyy-MM-dd'))" -Level "SUCCESS"
            $ActionsPerformed += "Account expiration set to $($ExpirationDate.ToString('yyyy-MM-dd'))"
        }
    }
    catch {
        $ErrorMessage = "Failed to set account expiration: $($_.Exception.Message)"
        Write-Log -Message $ErrorMessage -Level "ERROR"
        Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    }
}

# ============================================================================
# STEP 5: MOVE TO DISABLED OU
# ============================================================================

if ($DisabledOU) {
    Write-Log -Message "Step 5: Moving user to Disabled OU..." -Level "INFO"
    
    try {
        # Validate OU exists
        $null = Get-ADOrganizationalUnit -Identity $DisabledOU -ErrorAction Stop
        
        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Move to $DisabledOU")) {
            Move-ADObject -Identity $ADUser.DistinguishedName -TargetPath $DisabledOU -ErrorAction Stop
            Write-Log -Message "User moved to Disabled OU successfully" -Level "SUCCESS"
            $ActionsPerformed += "Moved to Disabled Users OU"
        }
    }
    catch {
        $ErrorMessage = "Failed to move user to Disabled OU: $($_.Exception.Message)"
        Write-Log -Message $ErrorMessage -Level "ERROR"
        Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    }
}

# ============================================================================
# STEP 6: UPDATE DESCRIPTION
# ============================================================================

Write-Log -Message "Step 6: Updating account description..." -Level "INFO"

try {
    $OffboardingNote = "Offboarded on $(Get-Date -Format 'yyyy-MM-dd') by $env:USERNAME"
    
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Update description")) {
        Set-ADUser -Identity $ADUser.DistinguishedName -Description $OffboardingNote -ErrorAction Stop
        Write-Log -Message "Account description updated" -Level "SUCCESS"
        $ActionsPerformed += "Description updated with offboarding date"
    }
}
catch {
    Write-Log -Message "Failed to update description: $($_.Exception.Message)" -Level "WARNING"
}

# ============================================================================
# STEP 7: HIDE FROM GAL (if Exchange attributes present)
# ============================================================================

if ($HideFromGAL) {
    Write-Log -Message "Step 7: Hiding from Global Address List..." -Level "INFO"
    
    try {
        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Hide from GAL")) {
            Set-ADUser -Identity $ADUser.DistinguishedName -Replace @{msExchHideFromAddressLists = $true} -ErrorAction Stop
            Write-Log -Message "User hidden from GAL" -Level "SUCCESS"
            $ActionsPerformed += "Hidden from Global Address List"
        }
    }
    catch {
        Write-Log -Message "Failed to hide from GAL: $($_.Exception.Message)" -Level "WARNING"
    }
}

# ============================================================================
# STEP 8: NOTIFY MANAGER
# ============================================================================

if ($NotifyManager -and $ADUser.Manager) {
    Write-Log -Message "Step 8: Notifying manager..." -Level "INFO"
    
    try {
        $Manager = Get-ADUser -Identity $ADUser.Manager -Properties EmailAddress -ErrorAction Stop
        
        if ($Manager.EmailAddress) {
            Send-ManagerNotification -ManagerEmail $Manager.EmailAddress `
                                      -UserDisplayName $ADUser.DisplayName `
                                      -UserUPN $UserPrincipalName `
                                      -ActionsPerformed $ActionsPerformed
        }
        else {
            Write-Log -Message "Manager found but no email address configured" -Level "WARNING"
        }
    }
    catch {
        Write-Log -Message "Failed to notify manager: $($_.Exception.Message)" -Level "WARNING"
    }
}

# ============================================================================
# SUMMARY REPORT
# ============================================================================

Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "User Offboarding Process Completed" -Level "INFO"
Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "User: $($ADUser.DisplayName) ($UserPrincipalName)" -Level "INFO"
Write-Log -Message "Actions Performed:" -Level "INFO"

foreach ($Action in $ActionsPerformed) {
    Write-Log -Message "  - $Action" -Level "SUCCESS"
}

Write-Log -Message "Log File: $LogFile" -Level "INFO"

# Return summary object
[PSCustomObject]@{
    UserPrincipalName = $UserPrincipalName
    DisplayName       = $ADUser.DisplayName
    SamAccountName    = $ADUser.SamAccountName
    OffboardingDate   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    ActionsPerformed  = $ActionsPerformed
    Success           = ($ActionsPerformed.Count -gt 0)
    LogFile           = $LogFile
}