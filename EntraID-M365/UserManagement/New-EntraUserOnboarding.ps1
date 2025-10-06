<#
.SYNOPSIS
    Comprehensive Entra ID (Azure AD) user onboarding automation script.

.DESCRIPTION
    This script automates the complete onboarding process for new cloud users including:
    - Creating Entra ID user account
    - Assigning initial security groups
    - Setting temporary password
    - Provisioning Exchange Online mailbox
    - Assigning Microsoft 365 licenses
    - Sending welcome email with credentials

.PARAMETER FirstName
    User's first name.

.PARAMETER LastName
    User's last name.

.PARAMETER DisplayName
    User's display name. If not provided, will be generated from FirstName and LastName.

.PARAMETER UserPrincipalName
    User Principal Name (email address) for the new user.

.PARAMETER Department
    Department name.

.PARAMETER JobTitle
    Job title for the user.

.PARAMETER SecurityGroups
    Array of security group names to add the user to.

.PARAMETER LicenseSKU
    Microsoft 365 license SKU to assign (e.g., 'ENTERPRISEPACK' for E3).

.PARAMETER ManagerUPN
    UPN of the user's manager.

.PARAMETER SendWelcomeEmail
    If specified, sends welcome email to the new user.

.PARAMETER UsageLocation
    Two-letter country code for license assignment. Default: US

.EXAMPLE
    .\New-EntraUserOnboarding.ps1 -FirstName "John" -LastName "Doe" -UserPrincipalName "john.doe@contoso.com" -Department "IT" -JobTitle "Systems Administrator" -LicenseSKU "ENTERPRISEPACK"

.EXAMPLE
    .\New-EntraUserOnboarding.ps1 -FirstName "Jane" -LastName "Smith" -UserPrincipalName "jane.smith@contoso.com" -SecurityGroups @("Sales Team", "VPN Users") -SendWelcomeEmail -Verbose

.NOTES
    Author: Tendai Choruwa
    Version: 1.0
    Last Updated: January 2025
    Requires: Microsoft.Graph PowerShell Module, Global Administrator or User Administrator role
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$FirstName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$LastName,

    [Parameter(Mandatory = $false)]
    [string]$DisplayName,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$Department,

    [Parameter(Mandatory = $false)]
    [string]$JobTitle,

    [Parameter(Mandatory = $false)]
    [string[]]$SecurityGroups,

    [Parameter(Mandatory = $false)]
    [string]$LicenseSKU,

    [Parameter(Mandatory = $false)]
    [string]$ManagerUPN,

    [Parameter(Mandatory = $false)]
    [switch]$SendWelcomeEmail,

    [Parameter(Mandatory = $false)]
    [ValidateLength(2, 2)]
    [string]$UsageLocation = "US"
)

#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Identity.DirectoryManagement

# ============================================================================
# INITIALIZATION
# ============================================================================

$ErrorActionPreference = "Stop"

# Create logs directory
$LogDirectory = "C:\Logs"
if (-not (Test-Path -Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

$LogFile = Join-Path -Path $LogDirectory -ChildPath "EntraUserOnboarding_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
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

function New-TempPassword {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [int]$Length = 16
    )

    $lowercase = "abcdefghjkmnpqrstuvwxyz"  # Removed ambiguous characters
    $uppercase = "ABCDEFGHJKLMNPQRSTUVWXYZ"
    $numbers = "23456789"  # Removed 0 and 1
    $special = "!@#$%^&*"
    
    $password = ""
    $password += $lowercase[(Get-Random -Maximum $lowercase.Length)]
    $password += $uppercase[(Get-Random -Maximum $uppercase.Length)]
    $password += $numbers[(Get-Random -Maximum $numbers.Length)]
    $password += $special[(Get-Random -Maximum $special.Length)]
    
    $allChars = $lowercase + $uppercase + $numbers + $special
    for ($i = 0; $i -lt ($Length - 4); $i++) {
        $password += $allChars[(Get-Random -Maximum $allChars.Length)]
    }
    
    # Shuffle
    $passwordArray = $password.ToCharArray()
    $shuffled = $passwordArray | Get-Random -Count $passwordArray.Length
    return (-join $shuffled)
}

function Send-WelcomeEmail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$RecipientUPN,
        
        [Parameter(Mandatory = $true)]
        [string]$UserDisplayName,
        
        [Parameter(Mandatory = $true)]
        [string]$TempPassword
    )

    try {
        $Subject = "Welcome to the Organization - Your Account is Ready!"
        
        $Body = @"
<html>
<head>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; text-align: center; }
        .content { padding: 30px; background-color: #f9f9f9; }
        .credentials { background-color: #fff; border-left: 4px solid #667eea; padding: 20px; margin: 20px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .important { background-color: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0; }
        .footer { text-align: center; padding: 20px; color: #666; font-size: 12px; }
        .button { background-color: #667eea; color: white; padding: 12px 30px; text-decoration: none; border-radius: 5px; display: inline-block; margin: 10px 0; }
        ul { list-style-type: none; padding-left: 0; }
        li { padding: 8px 0; padding-left: 25px; position: relative; }
        li:before { content: "✓"; position: absolute; left: 0; color: #667eea; font-weight: bold; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Welcome to Our Team!</h1>
        <p>We're excited to have you on board</p>
    </div>
    
    <div class="content">
        <h2>Hello $UserDisplayName,</h2>
        <p>Your user account has been created and is ready to use. Below are your login credentials and important information to get you started.</p>
        
        <div class="credentials">
            <h3>📧 Your Login Credentials</h3>
            <p><strong>Email/Username:</strong> $RecipientUPN</p>
            <p><strong>Temporary Password:</strong> <code style="background-color: #f0f0f0; padding: 5px 10px; border-radius: 3px; font-size: 16px;">$TempPassword</code></p>
        </div>
        
        <div class="important">
            <h3>⚠️ Important Security Notice</h3>
            <p><strong>You must change your password upon first login.</strong> Please choose a strong, unique password that you haven't used elsewhere.</p>
        </div>
        
        <h3>🚀 Getting Started</h3>
        <ul>
            <li>Access Microsoft 365 at: <a href="https://portal.office.com">https://portal.office.com</a></li>
            <li>Your mailbox will be available within 15 minutes</li>
            <li>OneDrive and Teams are ready for use</li>
            <li>Download Office apps from the portal</li>
        </ul>
        
        <h3>📱 Multi-Factor Authentication (MFA)</h3>
        <p>For security, you'll be prompted to set up MFA on your first login. Please have your mobile device ready.</p>
        
        <h3>📞 Need Help?</h3>
        <p>If you encounter any issues or have questions:</p>
        <ul>
            <li>Email: support@company.com</li>
            <li>Phone: +1 (555) 123-4567</li>
            <li>IT Portal: <a href="https://help.company.com">https://help.company.com</a></li>
        </ul>
        
        <center>
            <a href="https://portal.office.com" class="button">Access Microsoft 365</a>
        </center>
    </div>
    
    <div class="footer">
        <p>This is an automated message from IT Operations.</p>
        <p>Please do not reply to this email.</p>
        <p>© $(Get-Date -Format yyyy) Company Name. All rights reserved.</p>
    </div>
</body>
</html>
"@

        $MessageParams = @{
            Message         = @{
                Subject      = $Subject
                Body         = @{
                    ContentType = "HTML"
                    Content     = $Body
                }
                ToRecipients = @(
                    @{
                        EmailAddress = @{
                            Address = $RecipientUPN
                        }
                    }
                )
            }
            SaveToSentItems = $true
        }

        Send-MgUserMail -UserId "noreply@company.com" -BodyParameter $MessageParams -ErrorAction Stop
        Write-Log -Message "Welcome email sent to $RecipientUPN" -Level "SUCCESS"
    }
    catch {
        Write-Log -Message "Failed to send welcome email: $($_.Exception.Message)" -Level "WARNING"
    }
}

# ============================================================================
# MAIN SCRIPT EXECUTION
# ============================================================================

Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "Starting Entra ID User Onboarding" -Level "INFO"
Write-Log -Message "========================================" -Level "INFO"

# Set display name if not provided
if (-not $DisplayName) {
    $DisplayName = "$FirstName $LastName"
}

Write-Log -Message "User: $DisplayName ($UserPrincipalName)" -Level "INFO"

# Initialize actions list
$ActionsPerformed = @()

# ============================================================================
# STEP 1: CONNECT TO MICROSOFT GRAPH
# ============================================================================

Write-Log -Message "Step 1: Connecting to Microsoft Graph..." -Level "INFO"

try {
    # Check if already connected
    $context = Get-MgContext -ErrorAction SilentlyContinue
    
    if (-not $context) {
        Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All", "Directory.ReadWrite.All" -ErrorAction Stop
        Write-Log -Message "Connected to Microsoft Graph successfully" -Level "SUCCESS"
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
# STEP 2: CHECK IF USER EXISTS
# ============================================================================

Write-Log -Message "Step 2: Checking if user already exists..." -Level "INFO"

try {
    $ExistingUser = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction SilentlyContinue
    
    if ($ExistingUser) {
        $ErrorMessage = "User already exists: $UserPrincipalName"
        Write-Log -Message $ErrorMessage -Level "ERROR"
        Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
        exit 1
    }
    
    Write-Log -Message "User does not exist - proceeding with creation" -Level "SUCCESS"
}
catch {
    Write-Log -Message "Error checking user existence: $($_.Exception.Message)" -Level "WARNING"
}

# ============================================================================
# STEP 3: GENERATE TEMPORARY PASSWORD
# ============================================================================

Write-Log -Message "Step 3: Generating temporary password..." -Level "INFO"

$TempPassword = New-TempPassword
Write-Log -Message "Temporary password generated" -Level "SUCCESS"

# ============================================================================
# STEP 4: CREATE USER ACCOUNT
# ============================================================================

Write-Log -Message "Step 4: Creating Entra ID user account..." -Level "INFO"

$PasswordProfile = @{
    Password                             = $TempPassword
    ForceChangePasswordNextSignIn        = $true
    ForceChangePasswordNextSignInWithMfa = $true
}

$UserParams = @{
    AccountEnabled    = $true
    DisplayName       = $DisplayName
    GivenName         = $FirstName
    Surname           = $LastName
    UserPrincipalName = $UserPrincipalName
    MailNickname      = $UserPrincipalName.Split('@')[0]
    PasswordProfile   = $PasswordProfile
    UsageLocation     = $UsageLocation
}

# Add optional parameters
if ($Department) {
    $UserParams.Add("Department", $Department)
}
if ($JobTitle) {
    $UserParams.Add("JobTitle", $JobTitle)
}

try {
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Create Entra ID User")) {
        $NewUser = New-MgUser @UserParams -ErrorAction Stop
        Write-Log -Message "User created successfully with ID: $($NewUser.Id)" -Level "SUCCESS"
        $ActionsPerformed += "User account created in Entra ID"
    }
}
catch {
    $ErrorMessage = "Failed to create user: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - Stack Trace: $($_.ScriptStackTrace)"
    exit 1
}

# Wait for user object to replicate
Write-Log -Message "Waiting 10 seconds for user object replication..." -Level "INFO"
Start-Sleep -Seconds 10

# ============================================================================
# STEP 5: ASSIGN SECURITY GROUPS
# ============================================================================

if ($SecurityGroups -and $SecurityGroups.Count -gt 0) {
    Write-Log -Message "Step 5: Adding user to security groups..." -Level "INFO"
    
    foreach ($GroupName in $SecurityGroups) {
        try {
            $Group = Get-MgGroup -Filter "displayName eq '$GroupName'" -ErrorAction Stop
            
            if ($Group) {
                if ($PSCmdlet.ShouldProcess($GroupName, "Add user to group")) {
                    New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $NewUser.Id -ErrorAction Stop
                    Write-Log -Message "Added to group: $GroupName" -Level "SUCCESS"
                }
            }
            else {
                Write-Log -Message "Group not found: $GroupName" -Level "WARNING"
            }
        }
        catch {
            Write-Log -Message "Failed to add to group '$GroupName': $($_.Exception.Message)" -Level "WARNING"
        }
    }
    
    $ActionsPerformed += "Added to $($SecurityGroups.Count) security group(s)"
}

# ============================================================================
# STEP 6: ASSIGN MANAGER
# ============================================================================

if ($ManagerUPN) {
    Write-Log -Message "Step 6: Assigning manager..." -Level "INFO"
    
    try {
        $Manager = Get-MgUser -Filter "userPrincipalName eq '$ManagerUPN'" -ErrorAction Stop
        
        if ($Manager) {
            if ($PSCmdlet.ShouldProcess($ManagerUPN, "Assign as manager")) {
                Set-MgUserManagerByRef -UserId $NewUser.Id -BodyParameter @{"@odata.id" = "https://graph.microsoft.com/v1.0/users/$($Manager.Id)"} -ErrorAction Stop
                Write-Log -Message "Manager assigned: $($Manager.DisplayName)" -Level "SUCCESS"
                $ActionsPerformed += "Manager assigned: $($Manager.DisplayName)"
            }
        }
        else {
            Write-Log -Message "Manager not found: $ManagerUPN" -Level "WARNING"
        }
    }
    catch {
        Write-Log -Message "Failed to assign manager: $($_.Exception.Message)" -Level "WARNING"
    }
}

# ============================================================================
# STEP 7: ASSIGN LICENSE
# ============================================================================

if ($LicenseSKU) {
    Write-Log -Message "Step 7: Assigning Microsoft 365 license..." -Level "INFO"
    
    try {
        # Get available licenses
        $SubscribedSkus = Get-MgSubscribedSku -ErrorAction Stop
        $License = $SubscribedSkus | Where-Object { $_.SkuPartNumber -eq $LicenseSKU }
        
        if ($License) {
            if ($License.ConsumedUnits -lt $License.PrepaidUnits.Enabled) {
                if ($PSCmdlet.ShouldProcess($LicenseSKU, "Assign license")) {
                    $LicenseParams = @{
                        AddLicenses    = @(
                            @{
                                SkuId = $License.SkuId
                            }
                        )
                        RemoveLicenses = @()
                    }
                    
                    Set-MgUserLicense -UserId $NewUser.Id -BodyParameter $LicenseParams -ErrorAction Stop
                    Write-Log -Message "License assigned: $LicenseSKU" -Level "SUCCESS"
                    $ActionsPerformed += "License assigned: $LicenseSKU"
                }
            }
            else {
                Write-Log -Message "No available licenses for SKU: $LicenseSKU" -Level "WARNING"
            }
        }
        else {
            Write-Log -Message "License SKU not found: $LicenseSKU" -Level "WARNING"
        }
    }
    catch {
        $ErrorMessage = "Failed to assign license: $($_.Exception.Message)"
        Write-Log -Message $ErrorMessage -Level "WARNING"
        Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    }
}

# ============================================================================
# STEP 8: WAIT FOR MAILBOX PROVISIONING
# ============================================================================

if ($LicenseSKU) {
    Write-Log -Message "Step 8: Waiting for mailbox provisioning..." -Level "INFO"
    Write-Log -Message "Exchange Online mailbox will be available within 15 minutes" -Level "INFO"
    $ActionsPerformed += "Exchange Online mailbox provisioning initiated"
}

# ============================================================================
# STEP 9: SEND WELCOME EMAIL
# ============================================================================

if ($SendWelcomeEmail) {
    Write-Log -Message "Step 9: Sending welcome email..." -Level "INFO"
    
    # Wait a bit longer for mailbox
    Write-Log -Message "Waiting 30 seconds before sending welcome email..." -Level "INFO"
    Start-Sleep -Seconds 30
    
    Send-WelcomeEmail -RecipientUPN $UserPrincipalName -UserDisplayName $DisplayName -TempPassword $TempPassword
    $ActionsPerformed += "Welcome email sent"
}

# Clear sensitive data from memory
$TempPassword = $null
$PasswordProfile = $null

# ============================================================================
# SUMMARY REPORT
# ============================================================================

Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "User Onboarding Process Completed" -Level "INFO"
Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "User: $DisplayName" -Level "INFO"
Write-Log -Message "UPN: $UserPrincipalName" -Level "INFO"
Write-Log -Message "User ID: $($NewUser.Id)" -Level "INFO"
Write-Log -Message "" -Level "INFO"
Write-Log -Message "Actions Performed:" -Level "INFO"

foreach ($Action in $ActionsPerformed) {
    Write-Log -Message "  ✓ $Action" -Level "SUCCESS"
}

Write-Log -Message "" -Level "INFO"
Write-Log -Message "Next Steps:" -Level "INFO"
Write-Log -Message "  - User can sign in at https://portal.office.com" -Level "INFO"
Write-Log -Message "  - Password change required on first login" -Level "INFO"
Write-Log -Message "  - MFA setup required on first login" -Level "INFO"
Write-Log -Message "  - Mailbox available within 15 minutes" -Level "INFO"
Write-Log -Message "" -Level "INFO"
Write-Log -Message "Log File: $LogFile" -Level "INFO"

# Return summary object
[PSCustomObject]@{
    Success           = $true
    UserPrincipalName = $UserPrincipalName
    DisplayName       = $DisplayName
    UserId            = $NewUser.Id
    Department        = $Department
    JobTitle          = $JobTitle
    LicenseAssigned   = $LicenseSKU
    GroupsAdded       = $SecurityGroups
    ActionsPerformed  = $ActionsPerformed
    CreatedDate       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    LogFile           = $LogFile
}