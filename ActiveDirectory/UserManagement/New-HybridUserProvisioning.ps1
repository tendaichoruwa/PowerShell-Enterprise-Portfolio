<#
.SYNOPSIS
    Advanced hybrid user provisioning script for Active Directory with Entra ID sync.

.DESCRIPTION
    This script automates the creation of on-premises Active Directory user accounts with 
    properties configured for Azure AD Connect synchronization. It reads user details from 
    a CSV file, validates input data, creates AD user objects with standardized attributes,
    and prepares them for hybrid identity scenarios.

.PARAMETER CSVPath
    Path to the CSV file containing user information. Required columns:
    FirstName, LastName, Username, Email, Department, Title, Office, ManagerEmail

.PARAMETER TargetOU
    Distinguished Name of the target Organizational Unit for new users.

.PARAMETER DefaultPassword
    Secure string containing the default password. If not provided, a random password is generated.

.PARAMETER SendWelcomeEmail
    If specified, sends a welcome email to each newly created user.

.PARAMETER SMTPServer
    SMTP server address for sending welcome emails. Default: smtp.company.com

.PARAMETER EmailFrom
    Sender email address for welcome emails. Default: noreply@company.com

.EXAMPLE
    .\New-HybridUserProvisioning.ps1 -CSVPath "C:\Users\Import.csv" -TargetOU "OU=Users,DC=contoso,DC=com" -Verbose

.EXAMPLE
    .\New-HybridUserProvisioning.ps1 -CSVPath "C:\Users\Import.csv" -TargetOU "OU=Users,DC=contoso,DC=com" -SendWelcomeEmail -SMTPServer "mail.contoso.com"

.NOTES
    Author: Tendai Choruwa
    Version: 1.0
    Last Updated: January 2025
    Requires: Active Directory PowerShell Module, Domain Admin privileges
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$CSVPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TargetOU,

    [Parameter(Mandatory = $false)]
    [SecureString]$DefaultPassword,

    [Parameter(Mandatory = $false)]
    [switch]$SendWelcomeEmail,

    [Parameter(Mandatory = $false)]
    [string]$SMTPServer = "smtp.company.com",

    [Parameter(Mandatory = $false)]
    [string]$EmailFrom = "noreply@company.com"
)

#Requires -Modules ActiveDirectory

# ============================================================================
# INITIALIZATION
# ============================================================================

# Set error action preference
$ErrorActionPreference = "Stop"

# Create logs directory if it doesn't exist
$LogDirectory = "C:\Logs"
if (-not (Test-Path -Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

# Set log file paths
$LogFile = Join-Path -Path $LogDirectory -ChildPath "UserProvisioning_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$ErrorLogFile = Join-Path -Path $LogDirectory -ChildPath "Errors.txt"

# ============================================================================
# FUNCTIONS
# ============================================================================

function Write-Log {
    <#
    .SYNOPSIS
        Writes a message to the log file and console.
    #>
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
    
    # Write to log file
    Add-Content -Path $LogFile -Value $LogEntry
    
    # Write to console with color
    switch ($Level) {
        "INFO"    { Write-Host $LogEntry -ForegroundColor Cyan }
        "SUCCESS" { Write-Host $LogEntry -ForegroundColor Green }
        "WARNING" { Write-Host $LogEntry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $LogEntry -ForegroundColor Red }
    }
}

function New-RandomPassword {
    <#
    .SYNOPSIS
        Generates a cryptographically secure random password.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [int]$Length = 16
    )

    $lowercase = "abcdefghijklmnopqrstuvwxyz"
    $uppercase = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $numbers = "0123456789"
    $special = "!@#$%^&*"
    
    $allChars = $lowercase + $uppercase + $numbers + $special
    
    $password = -join (
        $lowercase[(Get-Random -Maximum $lowercase.Length)],
        $uppercase[(Get-Random -Maximum $uppercase.Length)],
        $numbers[(Get-Random -Maximum $numbers.Length)],
        $special[(Get-Random -Maximum $special.Length)]
    )
    
    for ($i = 0; $i -lt ($Length - 4); $i++) {
        $password += $allChars[(Get-Random -Maximum $allChars.Length)]
    }
    
    # Shuffle the password
    $passwordArray = $password.ToCharArray()
    $shuffled = $passwordArray | Get-Random -Count $passwordArray.Length
    return (-join $shuffled)
}

function Test-UserExists {
    <#
    .SYNOPSIS
        Checks if a user already exists in Active Directory.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$SamAccountName
    )

    try {
        $existingUser = Get-ADUser -Filter "SamAccountName -eq '$SamAccountName'" -ErrorAction SilentlyContinue
        return ($null -ne $existingUser)
    }
    catch {
        return $false
    }
}

function Send-WelcomeEmail {
    <#
    .SYNOPSIS
        Sends a welcome email to the newly created user.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ToAddress,
        
        [Parameter(Mandatory = $true)]
        [string]$Username,
        
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,
        
        [Parameter(Mandatory = $true)]
        [string]$TempPassword
    )

    try {
        $Subject = "Welcome to the Organization - Your Account Details"
        
        $Body = @"
<html>
<body style='font-family: Arial, sans-serif;'>
    <h2>Welcome to the Organization, $DisplayName!</h2>
    <p>Your user account has been successfully created. Below are your login credentials:</p>
    
    <div style='background-color: #f5f5f5; padding: 15px; border-left: 4px solid #0066cc;'>
        <p><strong>Username:</strong> $Username</p>
        <p><strong>Temporary Password:</strong> $TempPassword</p>
        <p><strong>Email:</strong> $ToAddress</p>
    </div>
    
    <p><strong style='color: #cc0000;'>IMPORTANT:</strong> You will be required to change your password upon first login.</p>
    
    <h3>Getting Started:</h3>
    <ul>
        <li>Access company portal at: <a href='https://portal.company.com'>https://portal.company.com</a></li>
        <li>Your account will sync to cloud services within 30 minutes</li>
        <li>Contact IT Support for any issues: support@company.com</li>
    </ul>
    
    <p>If you have any questions, please don't hesitate to reach out to the IT Help Desk.</p>
    
    <hr>
    <p style='font-size: 12px; color: #666;'>This is an automated message. Please do not reply to this email.</p>
</body>
</html>
"@

        $EmailParams = @{
            To         = $ToAddress
            From       = $EmailFrom
            Subject    = $Subject
            Body       = $Body
            BodyAsHtml = $true
            SmtpServer = $SMTPServer
            Priority   = "High"
        }

        Send-MailMessage @EmailParams -ErrorAction Stop
        Write-Log -Message "Welcome email sent to $ToAddress" -Level "SUCCESS"
    }
    catch {
        Write-Log -Message "Failed to send welcome email to $ToAddress : $($_.Exception.Message)" -Level "WARNING"
    }
}

# ============================================================================
# MAIN SCRIPT EXECUTION
# ============================================================================

Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "Starting Hybrid User Provisioning Process" -Level "INFO"
Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "CSV File: $CSVPath" -Level "INFO"
Write-Log -Message "Target OU: $TargetOU" -Level "INFO"

# Validate Active Directory module
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log -Message "Active Directory module loaded successfully" -Level "SUCCESS"
}
catch {
    Write-Log -Message "Failed to load Active Directory module: $($_.Exception.Message)" -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - Failed to load AD module: $($_.Exception.Message)"
    exit 1
}

# Validate OU exists
try {
    $null = Get-ADOrganizationalUnit -Identity $TargetOU -ErrorAction Stop
    Write-Log -Message "Target OU validated successfully" -Level "SUCCESS"
}
catch {
    Write-Log -Message "Target OU '$TargetOU' not found or inaccessible" -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - Invalid OU: $TargetOU"
    exit 1
}

# Import CSV file
try {
    $Users = Import-Csv -Path $CSVPath -ErrorAction Stop
    Write-Log -Message "Successfully imported $($Users.Count) users from CSV" -Level "SUCCESS"
}
catch {
    Write-Log -Message "Failed to import CSV file: $($_.Exception.Message)" -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - CSV Import Failed: $($_.Exception.Message)"
    exit 1
}

# Validate CSV columns
$RequiredColumns = @("FirstName", "LastName", "Username", "Email", "Department", "Title")
$CSVColumns = $Users[0].PSObject.Properties.Name

$MissingColumns = $RequiredColumns | Where-Object { $_ -notin $CSVColumns }
if ($MissingColumns) {
    Write-Log -Message "CSV is missing required columns: $($MissingColumns -join ', ')" -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - Missing CSV columns: $($MissingColumns -join ', ')"
    exit 1
}

# Initialize counters
$SuccessCount = 0
$FailCount = 0
$SkipCount = 0

# Process each user
foreach ($User in $Users) {
    Write-Log -Message "----------------------------------------" -Level "INFO"
    Write-Log -Message "Processing user: $($User.Username)" -Level "INFO"

    # Validate required fields
    if ([string]::IsNullOrWhiteSpace($User.FirstName) -or 
        [string]::IsNullOrWhiteSpace($User.LastName) -or 
        [string]::IsNullOrWhiteSpace($User.Username) -or 
        [string]::IsNullOrWhiteSpace($User.Email)) {
        
        Write-Log -Message "Skipping user due to missing required fields" -Level "WARNING"
        $SkipCount++
        continue
    }

    # Check if user already exists
    if (Test-UserExists -SamAccountName $User.Username) {
        Write-Log -Message "User '$($User.Username)' already exists. Skipping." -Level "WARNING"
        $SkipCount++
        continue
    }

    # Generate or use provided password
    if ($DefaultPassword) {
        $SecurePassword = $DefaultPassword
        $PlainPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($DefaultPassword))
    }
    else {
        $PlainPassword = New-RandomPassword
        $SecurePassword = ConvertTo-SecureString -String $PlainPassword -AsPlainText -Force
    }

    # Construct user properties
    $DisplayName = "$($User.FirstName) $($User.LastName)"
    $UserPrincipalName = $User.Email

    $UserParams = @{
        SamAccountName        = $User.Username
        UserPrincipalName     = $UserPrincipalName
        Name                  = $DisplayName
        GivenName             = $User.FirstName
        Surname               = $User.LastName
        DisplayName           = $DisplayName
        EmailAddress          = $User.Email
        Department            = $User.Department
        Title                 = $User.Title
        Path                  = $TargetOU
        AccountPassword       = $SecurePassword
        Enabled               = $true
        ChangePasswordAtLogon = $true
        PasswordNeverExpires  = $false
        CannotChangePassword  = $false
    }

    # Add optional fields if present
    if (-not [string]::IsNullOrWhiteSpace($User.Office)) {
        $UserParams.Add("Office", $User.Office)
    }
    if (-not [string]::IsNullOrWhiteSpace($User.EmployeeID)) {
        $UserParams.Add("EmployeeID", $User.EmployeeID)
    }
    if (-not [string]::IsNull