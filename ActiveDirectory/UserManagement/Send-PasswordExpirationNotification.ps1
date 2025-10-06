[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 90)]
    [int]$DaysBeforeExpiration = 14,

    [Parameter(Mandatory = $false)]
    [switch]$NotifyManager,

    [Parameter(Mandatory = $false)]
    [string]$SMTPServer = "smtp.company.com",

    [Parameter(Mandatory = $false)]
    [string]$EmailFrom = "noreply@company.com",

    [Parameter(Mandatory = $false)]
    [string[]]$ExcludeOUs,

    [Parameter(Mandatory = $false)]
    [switch]$TestMode,

    [Parameter(Mandatory = $false)]
    [switch]$GenerateReport,

    [Parameter(Mandatory = $false)]
    [string]$ReportPath
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

$LogFile = Join-Path -Path $LogDirectory -ChildPath "PasswordExpiration_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$ErrorLogFile = Join-Path -Path $LogDirectory -ChildPath "Errors.txt"

if (-not $ReportPath) {
    $ReportPath = Join-Path -Path $LogDirectory -ChildPath "PasswordExpiration_Report_$(Get-Date -Format 'yyyyMMdd').html"
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
        "INFO" { Write-Host $LogEntry -ForegroundColor Cyan }
        "SUCCESS" { Write-Host $LogEntry -ForegroundColor Green }
        "WARNING" { Write-Host $LogEntry -ForegroundColor Yellow }
        "ERROR" { Write-Host $LogEntry -ForegroundColor Red }
    }
}

function Get-PasswordExpirationInfo {
    <#
    .SYNOPSIS
        Calculates password expiration details for a user.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.ActiveDirectory.Management.ADUser]$User,
        
        [Parameter(Mandatory = $true)]
        [int]$MaxPasswordAge
    )

    if ($User.PasswordNeverExpires -or $null -eq $User.PasswordLastSet) {
        return $null
    }

    $ExpirationDate = $User.PasswordLastSet.AddDays($MaxPasswordAge)
    $DaysUntilExpiration = ($ExpirationDate - (Get-Date)).Days

    return [PSCustomObject]@{
        SamAccountName         = $User.SamAccountName
        DisplayName            = $User.DisplayName
        EmailAddress           = $User.EmailAddress
        PasswordLastSet        = $User.PasswordLastSet
        PasswordExpirationDate = $ExpirationDate
        DaysUntilExpiration    = $DaysUntilExpiration
        Manager                = $User.Manager
        Department             = $User.Department
        Title                  = $User.Title
    }
}

function Send-PasswordExpirationEmail {
    <#
    .SYNOPSIS
        Sends a branded password expiration notification email.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory = $false)]
        [string]$ManagerEmail
    )

    try {
        $DaysText = if ($UserInfo.DaysUntilExpiration -eq 1) { "1 day" } else { "$($UserInfo.DaysUntilExpiration) days" }
        $UrgencyColor = if ($UserInfo.DaysUntilExpiration -le 3) { "#dc3545" } elseif ($UserInfo.DaysUntilExpiration -le 7) { "#ffc107" } else { "#17a2b8" }
        $UrgencyLabel = if ($UserInfo.DaysUntilExpiration -le 3) { "URGENT" } elseif ($UserInfo.DaysUntilExpiration -le 7) { "IMPORTANT" } else { "REMINDER" }

        $Subject = "[$UrgencyLabel] Your password will expire in $DaysText"

        $Body = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            margin: 0;
            padding: 0;
        }
        .container {
            max-width: 600px;
            margin: 0 auto;
            background-color: #ffffff;
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 {
            margin: 0;
            font-size: 28px;
        }
        .urgency-banner {
            background-color: $UrgencyColor;
            color: white;
            padding: 20px;
            text-align: center;
            font-size: 22px;
            font-weight: bold;
        }
        .content {
            padding: 30px;
            background-color: #f9f9f9;
        }
        .info-box {
            background-color: white;
            border-left: 4px solid #667eea;
            padding: 20px;
            margin: 20px 0;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        .info-box h3 {
            margin-top: 0;
            color: #667eea;
        }
        .expiration-date {
            font-size: 24px;
            font-weight: bold;
            color: $UrgencyColor;
            text-align: center;
            padding: 15px;
            background-color: #fff;
            border-radius: 8px;
            margin: 20px 0;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        .instructions {
            background-color: #e7f3ff;
            border-left: 4px solid #0066cc;
            padding: 20px;
            margin: 20px 0;
        }
        .instructions h3 {
            margin-top: 0;
            color: #0066cc;
        }
        .instructions ol {
            padding-left: 20px;
        }
        .instructions li {
            padding: 8px 0;
        }
        .tips {
            background-color: #fff3cd;
            border-left: 4px solid #ffc107;
            padding: 20px;
            margin: 20px 0;
        }
        .tips h3 {
            margin-top: 0;
            color: #856404;
        }
        .tips ul {
            padding-left: 20px;
        }
        .tips li {
            padding: 5px 0;
        }
        .button {
            display: inline-block;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 40px;
            text-decoration: none;
            border-radius: 5px;
            font-weight: bold;
            text-align: center;
            margin: 20px 0;
        }
        .footer {
            background-color: #343a40;
            color: white;
            padding: 20px;
            text-align: center;
            font-size: 12px;
        }
        .footer p {
            margin: 5px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔐 Password Expiration Notice</h1>
        </div>
        
        <div class="urgency-banner">
            ⚠️ ${UrgencyLabel}: Password Expires in $DaysText
        </div>
        
        <div class="content">
            <p>Hello <strong>$($UserInfo.DisplayName)</strong>,</p>
            
            <p>This is a friendly reminder that your network password will expire soon. Please change your password before it expires to avoid any disruption to your work.</p>
            
            <div class="expiration-date">
                📅 Expiration Date: $($UserInfo.PasswordExpirationDate.ToString("dddd, MMMM dd, yyyy"))
                <br>
                ⏰ Time Remaining: <span style="color: $UrgencyColor;">$DaysText</span>
            </div>
            
            <div class="info-box">
                <h3>📋 Your Account Information</h3>
                <p><strong>Username:</strong> $($UserInfo.SamAccountName)</p>
                <p><strong>Email:</strong> $($UserInfo.EmailAddress)</p>
                <p><strong>Department:</strong> $($UserInfo.Department)</p>
                <p><strong>Last Password Change:</strong> $($UserInfo.PasswordLastSet.ToString("yyyy-MM-dd"))</p>
            </div>
            
            <div class="instructions">
                <h3>🔄 How to Change Your Password</h3>
                <p><strong>Option 1: On Windows Computer</strong></p>
                <ol>
                    <li>Press <strong>Ctrl + Alt + Del</strong></li>
                    <li>Click <strong>"Change a password"</strong></li>
                    <li>Enter your old password</li>
                    <li>Enter and confirm your new password</li>
                    <li>Click OK</li>
                </ol>
                
                <p><strong>Option 2: Online Portal</strong></p>
                <ol>
                    <li>Visit: <a href="https://passwordreset.microsoftonline.com">https://passwordreset.microsoftonline.com</a></li>
                    <li>Sign in with your credentials</li>
                    <li>Follow the password change wizard</li>
                </ol>
            </div>
            
            <div class="tips">
                <h3>💡 Password Requirements & Tips</h3>
                <ul>
                    <li>Minimum 12 characters long</li>
                    <li>Must contain uppercase and lowercase letters</li>
                    <li>Must contain at least one number</li>
                    <li>Must contain at least one special character (!@#$%^&*)</li>
                    <li>Cannot be the same as your previous 24 passwords</li>
                    <li>Cannot contain your username or common words</li>
                </ul>
                
                <p><strong>Strong Password Examples:</strong></p>
                <ul>
                    <li>Use a passphrase: <em>Coffee@Morning2025!</em></li>
                    <li>Use a sentence: <em>ILove2Travel!2025</em></li>
                    <li>Use a password manager for complex passwords</li>
                </ul>
            </div>
            
            <center>
                <a href="https://passwordreset.microsoftonline.com" class="button">Change Password Now</a>
            </center>
            
            <div class="info-box">
                <h3>❓ Need Help?</h3>
                <p>If you have any questions or need assistance changing your password:</p>
                <ul>
                    <li>📧 Email: <a href="mailto:support@company.com">support@company.com</a></li>
                    <li>📞 Phone: +1 (555) 123-4567</li>
                    <li>💬 IT Portal: <a href="https://help.company.com">https://help.company.com</a></li>
                </ul>
            </div>
            
            <p style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; color: #666; font-size: 14px;">
                <strong>What happens if my password expires?</strong><br>
                If your password expires, you will be unable to access company resources including email, file shares, and applications. Please change your password as soon as possible to avoid disruption.
            </p>
        </div>
        
        <div class="footer">
            <p><strong>IT Security Team</strong></p>
            <p>This is an automated notification. Please do not reply to this email.</p>
            <p>© $(Get-Date -Format yyyy) Company Name. All rights reserved.</p>
        </div>
    </div>
</body>
</html>
"@

        $Recipients = @($UserInfo.EmailAddress)
        
        if ($NotifyManager -and $ManagerEmail) {
            $Recipients += $ManagerEmail
        }

        $EmailParams = @{
            To         = $Recipients
            From       = $EmailFrom
            Subject    = $Subject
            Body       = $Body
            BodyAsHtml = $true
            SmtpServer = $SMTPServer
            Priority   = if ($UserInfo.DaysUntilExpiration -le 3) { "High" } else { "Normal" }
        }

        if ($TestMode) {
            Write-Log -Message "[TEST MODE] Would send email to: $($Recipients -join ', ')" -Level "INFO"
            Write-Log -Message "[TEST MODE] Subject: $Subject" -Level "INFO"
            return $true
        }
        else {
            Send-MailMessage @EmailParams -ErrorAction Stop
            Write-Log -Message "Email sent to: $($UserInfo.EmailAddress)" -Level "SUCCESS"
            return $true
        }
    }
    catch {
        $ErrorMessage = "Failed to send email to $($UserInfo.EmailAddress): $($_.Exception.Message)"
        Write-Log -Message $ErrorMessage -Level "ERROR"
        Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
        return $false
    }
}

function New-ExpirationReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array]$ExpiringUsers,
        
        [Parameter(Mandatory = $true)]
        [int]$DaysThreshold,
        
        [Parameter(Mandatory = $true)]
        [int]$NotificationsSent
    )

    $userRows = ""
    
    foreach ($user in $ExpiringUsers | Sort-Object DaysUntilExpiration) {
        $urgencyColor = if ($user.DaysUntilExpiration -le 3) { "#f8d7da" } 
        elseif ($user.DaysUntilExpiration -le 7) { "#fff3cd" } 
        else { "#d1ecf1" }
        
        $urgencyBadge = if ($user.DaysUntilExpiration -le 3) {
            "<span style='background-color: #dc3545; color: white; padding: 4px 8px; border-radius: 4px;'>🔴 URGENT</span>"
        }
        elseif ($user.DaysUntilExpiration -le 7) {
            "<span style='background-color: #ffc107; color: black; padding: 4px 8px; border-radius: 4px;'>⚠️ IMPORTANT</span>"
        }
        else {
            "<span style='background-color: #17a2b8; color: white; padding: 4px 8px; border-radius: 4px;'>ℹ️ REMINDER</span>"
        }
        
        $notificationStatus = if ($user.NotificationSent) {
            "<span style='color: #28a745;'>✓ Sent</span>"
        }
        else {
            "<span style='color: #dc3545;'>✗ Failed</span>"
        }
        
        $userRows += @"
        <tr style="background-color: $urgencyColor;">
            <td>$($user.SamAccountName)</td>
            <td>$($user.DisplayName)</td>
            <td>$($user.EmailAddress)</td>
            <td>$($user.Department)</td>
            <td>$($user.PasswordLastSet.ToString("yyyy-MM-dd"))</td>
            <td>$($user.PasswordExpirationDate.ToString("yyyy-MM-dd"))</td>
            <td style="text-align: center; font-weight: bold;">$($user.DaysUntilExpiration)</td>
            <td style="text-align: center;">$urgencyBadge</td>
            <td style="text-align: center;">$notificationStatus</td>
        </tr>
"@
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Password Expiration Notification Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            color: #333;
        }
        .container {
            max-width: 1600px;
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
            font-size: 14px;
        }
        th {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 12px 10px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 11px;
        }
        td {
            padding: 10px;
            border-bottom: 1px solid #dee2e6;
        }
        tr:hover {
            background-color: #f8f9fa !important;
        }
        .footer {
            background-color: #343a40;
            color: white;
            padding: 20px;
            text-align: center;
            font-size: 14px;
        }
        .alert {
            background-color: #d1ecf1;
            border-left: 4px solid #0c5460;
            padding: 15px;
            margin: 20px 0;
            border-radius: 4px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔐 Password Expiration Notification Report</h1>
            <p>Generated on $(Get-Date -Format "MMMM dd, yyyy 'at' HH:mm:ss")</p>
        </div>
        
        <div class="summary">
            <div class="summary-card">
                <h3>Users Expiring Soon</h3>
                <div class="value">$($ExpiringUsers.Count)</div>
            </div>
            <div class="summary-card">
                <h3>Notifications Sent</h3>
                <div class="value" style="color: #28a745;">$NotificationsSent</div>
            </div>
            <div class="summary-card">
                <h3>Urgent (≤3 days)</h3>
                <div class="value" style="color: #dc3545;">$(($ExpiringUsers | Where-Object { $_.DaysUntilExpiration -le 3 }).Count)</div>
            </div>
            <div class="summary-card">
                <h3>Warning Threshold</h3>
                <div class="value" style="font-size: 24px;">$DaysThreshold days</div>
            </div>
        </div>
        
        <div class="section">
            <div class="alert">
                <strong>ℹ️ Report Summary:</strong> This report shows all users whose passwords will expire within $DaysThreshold days. 
                Notifications have been sent to users to prompt them to change their passwords.
            </div>
        </div>
        
        <div class="section">
            <h2>📊 Expiring User Accounts</h2>
            <table>
                <thead>
                    <tr>
                        <th>Username</th>
                        <th>Display Name</th>
                        <th>Email</th>
                        <th>Department</th>
                        <th>Last Changed</th>
                        <th>Expires On</th>
                        <th style="text-align: center;">Days Left</th>
                        <th style="text-align: center;">Priority</th>
                        <th style="text-align: center;">Notification</th>
                    </tr>
                </thead>
                <tbody>
                    $userRows
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <p><strong>Password Expiration Notification System</strong></p>
            <p>Automated daily notifications ensure users maintain secure access</p>
            <p>Log file: $LogFile</p>
            <p>© $(Get-Date -Format yyyy) IT Operations</p>
        </div>
    </div>
</body>
</html>
"@

    return $html
}

# ============================================================================
# MAIN SCRIPT EXECUTION
# ============================================================================

Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "Password Expiration Notification System" -Level "INFO"
Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "Warning Threshold: $DaysBeforeExpiration days" -Level "INFO"

if ($TestMode) {
    Write-Log -Message "🧪 RUNNING IN TEST MODE - No emails will be sent" -Level "WARNING"
}

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

# ============================================================================
# STEP 1: GET DOMAIN PASSWORD POLICY
# ============================================================================

Write-Log -Message "Step 1: Retrieving domain password policy..." -Level "INFO"

try {
    $PasswordPolicy = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop
    $MaxPasswordAge = $PasswordPolicy.MaxPasswordAge.Days
    
    Write-Log -Message "Maximum password age: $MaxPasswordAge days" -Level "INFO"
    
    if ($MaxPasswordAge -eq 0) {
        Write-Log -Message "Password expiration is disabled in domain policy" -Level "WARNING"
        exit 0
    }
}
catch {
    $ErrorMessage = "Failed to retrieve password policy: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    exit 1
}

# ============================================================================
# STEP 2: GET ALL ENABLED USERS
# ============================================================================

Write-Log -Message "Step 2: Retrieving enabled user accounts..." -Level "INFO"

try {
    $FilterScript = {
        Enabled -eq $true -and 
        PasswordNeverExpires -eq $false -and 
        PasswordLastSet -ne $null
    }
    
    $Properties = @(
        'DisplayName',
        'EmailAddress',
        'PasswordLastSet',
        'PasswordNeverExpires',
        'Manager',
        'Department',
        'Title'
    )
    
    $AllUsers = Get-ADUser -Filter $FilterScript -Properties $Properties -ErrorAction Stop
    
    Write-Log -Message "Found $($AllUsers.Count) enabled users with expiring passwords" -Level "SUCCESS"
}
catch {
    $ErrorMessage = "Failed to retrieve users: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    exit 1
}

# Filter out excluded OUs
if ($ExcludeOUs) {
    Write-Log -Message "Filtering excluded OUs..." -Level "INFO"
    
    $AllUsers = $AllUsers | Where-Object {
        $userDN = $_.DistinguishedName
        $exclude = $false
        
        foreach ($ou in $ExcludeOUs) {
            if ($userDN -like "*$ou*") {
                $exclude = $true
                break
            }
        }
        
        -not $exclude
    }
    
    Write-Log -Message "After exclusions: $($AllUsers.Count) users" -Level "INFO"
}

# ============================================================================
# STEP 3: IDENTIFY USERS WITH EXPIRING PASSWORDS
# ============================================================================

Write-Log -Message "Step 3: Identifying users with passwords expiring in $DaysBeforeExpiration days..." -Level "INFO"

$ExpiringUsers = @()

foreach ($User in $AllUsers) {
    $ExpirationInfo = Get-PasswordExpirationInfo -User $User -MaxPasswordAge $MaxPasswordAge
    
    if ($ExpirationInfo -and 
        $ExpirationInfo.DaysUntilExpiration -le $DaysBeforeExpiration -and 
        $ExpirationInfo.DaysUntilExpiration -ge 0) {
        
        $ExpiringUsers += $ExpirationInfo
        
        $logLevel = if ($ExpirationInfo.DaysUntilExpiration -le 3) { "WARNING" } else { "INFO" }
        Write-Log -Message "  $($User.SamAccountName): $($ExpirationInfo.DaysUntilExpiration) days until expiration" -Level $logLevel
    }
}

Write-Log -Message "Found $($ExpiringUsers.Count) user(s) with passwords expiring within threshold" -Level $(if ($ExpiringUsers.Count -gt 0) { "WARNING" } else { "SUCCESS" })

# ============================================================================
# STEP 4: SEND NOTIFICATIONS
# ============================================================================

if ($ExpiringUsers.Count -gt 0) {
    Write-Log -Message "Step 4: Sending notifications..." -Level "INFO"
    
    $NotificationsSent = 0
    
    foreach ($UserInfo in $ExpiringUsers) {
        # Validate email address
        if ([string]::IsNullOrWhiteSpace($UserInfo.EmailAddress)) {
            Write-Log -Message "Skipping $($UserInfo.SamAccountName): No email address" -Level "WARNING"
            $UserInfo | Add-Member -NotePropertyName "NotificationSent" -NotePropertyValue $false -Force
            continue
        }
        
        # Get manager email if needed
        $ManagerEmail = $null
        if ($NotifyManager -and $UserInfo.Manager) {
            try {
                $Manager = Get-ADUser -Identity $UserInfo.Manager -Properties EmailAddress -ErrorAction Stop
                $ManagerEmail = $Manager.EmailAddress
            }
            catch {
                Write-Log -Message "Could not retrieve manager for $($UserInfo.SamAccountName)" -Level "WARNING"
            }
        }
        
        # Send notification
        $Sent = Send-PasswordExpirationEmail -UserInfo $UserInfo -ManagerEmail $ManagerEmail
        $UserInfo | Add-Member -NotePropertyName "NotificationSent" -NotePropertyValue $Sent -Force
        
        if ($Sent) {
            $NotificationsSent++
        }
        
        # Small delay to avoid overwhelming SMTP server
        Start-Sleep -Milliseconds 500
    }
    
    Write-Log -Message "Notifications sent: $NotificationsSent / $($ExpiringUsers.Count)" -Level "SUCCESS"
}
else {
    Write-Log -Message "No notifications to send" -Level "INFO"
    $NotificationsSent = 0
}

# ============================================================================
# STEP 5: GENERATE REPORT
# ============================================================================

if ($GenerateReport -or $ExpiringUsers.Count -gt 0) {
    Write-Log -Message "Step 5: Generating HTML report..." -Level "INFO"
    
    try {
        $htmlReport = New-ExpirationReport -ExpiringUsers $ExpiringUsers `
            -DaysThreshold $DaysBeforeExpiration `
            -NotificationsSent $NotificationsSent
        
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
# SUMMARY REPORT
# ============================================================================

Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "Notification Process Completed" -Level "INFO"
Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "Total users checked: $($AllUsers.Count)" -Level "INFO"
Write-Log -Message "Users with expiring passwords: $($ExpiringUsers.Count)" -Level $(if ($ExpiringUsers.Count -gt 0) { "WARNING" } else { "SUCCESS" })
Write-Log -Message "Notifications sent: $NotificationsSent" -Level "SUCCESS"

if ($ExpiringUsers.Count -gt 0) {
    Write-Log -Message "" -Level "INFO"
    Write-Log -Message "Breakdown by urgency:" -Level "INFO"
    
    $Urgent = ($ExpiringUsers | Where-Object { $_.DaysUntilExpiration -le 3 }).Count
    $Important = ($ExpiringUsers | Where-Object { $_.DaysUntilExpiration -gt 3 -and $_.DaysUntilExpiration -le 7 }).Count
    $Reminder = ($ExpiringUsers | Where-Object { $_.DaysUntilExpiration -gt 7 }).Count
    
    Write-Log -Message "  🔴 Urgent (≤3 days): $Urgent" -Level $(if ($Urgent -gt 0) { "ERROR" } else { "INFO" })
    Write-Log -Message "  ⚠️  Important (4-7 days): $Important" -Level $(if ($Important -gt 0) { "WARNING" } else { "INFO" })
    Write-Log -Message "  ℹ️  Reminder (8+ days): $Reminder" -Level "INFO"
}

Write-Log -Message "" -Level "INFO"
Write-Log -Message "Log File: $LogFile" -Level "INFO"

if ($GenerateReport -or $ExpiringUsers.Count -gt 0) {
    Write-Log -Message "Report: $ReportPath" -Level "INFO"
}

if ($TestMode) {
    Write-Log -Message "" -Level "INFO"
    Write-Log -Message "🧪 TEST MODE - No actual emails were sent" -Level "WARNING"
}

# Return summary object
[PSCustomObject]@{
    ExecutionDate          = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    DaysBeforeExpiration   = $DaysBeforeExpiration
    TotalUsersChecked      = $AllUsers.Count
    ExpiringPasswordsFound = $ExpiringUsers.Count
    NotificationsSent      = $NotificationsSent
    UrgentCount            = ($ExpiringUsers | Where-Object { $_.DaysUntilExpiration -le 3 }).Count
    ImportantCount         = ($ExpiringUsers | Where-Object { $_.DaysUntilExpiration -gt 3 -and $_.DaysUntilExpiration -le 7 }).Count
    ReminderCount          = ($ExpiringUsers | Where-Object { $_.DaysUntilExpiration -gt 7 }).Count
    TestMode               = $TestMode.IsPresent
    ReportPath             = $ReportPath
    LogFile                = $LogFile
    ExpiringUsers          = $ExpiringUsers
}<#
.SYNOPSIS
    Automated password expiration notification system for Active Directory users.

.DESCRIPTION
    This script identifies users whose passwords will expire soon and sends customized
    email notifications to users and optionally their managers. Features include:
    - Configurable expiration warning thresholds
    - Branded HTML email notifications
    - Manager CC notifications
    - Exclusion of service accounts and disabled users
    - Detailed logging and reporting
    - HTML summary report generation

.PARAMETER DaysBeforeExpiration
    Number of days before expiration to send notifications. Default: 14 days.

.PARAMETER NotifyManager
    If specified, CC's the user's manager on the notification email.

.PARAMETER SMTPServer
    SMTP server address for sending emails. Default: smtp.company.com

.PARAMETER EmailFrom
    Sender email address. Default: noreply@company.com

.PARAMETER ExcludeOUs
    Array of OU Distinguished Names to exclude from notifications (e.g., service accounts).

.PARAMETER TestMode
    If specified, displays what would be sent without actually sending emails.

.PARAMETER GenerateReport
    If specified, generates an HTML summary report of all notifications.

.PARAMETER ReportPath
    Path to save the HTML report. Default: C:\Logs\PasswordExpiration_Report_YYYYMMDD.html

.EXAMPLE
    .\Send-PasswordExpirationNotification.ps1 -DaysBeforeExpiration 14 -Verbose

.EXAMPLE
    .\Send-PasswordExpirationNotification.ps1 -DaysBeforeExpiration 7 -NotifyManager -SMTPServer "mail.contoso.com"

.EXAMPLE
    .\Send-PasswordExpirationNotification.ps1 -TestMode -GenerateReport

.NOTES
    Author: Tendai Choruwa
    Version: 1.0
    Last Updated: January 2025
    Requires: Active Directory PowerShell Module
    Recommended: Schedule as daily task at 8:00 AM
#>
