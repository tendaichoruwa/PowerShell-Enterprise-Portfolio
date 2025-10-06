# 🔐 Locked Accounts & Password Expiration Scripts - Usage Guide

## Overview

Two powerful scripts for Active Directory account security and password management:

1. **Find-LockedAccounts.ps1** - Detect and unlock locked accounts
2. **Send-PasswordExpirationNotification.ps1** - Automated password expiry notifications

---

## 📋 Find-LockedAccounts.ps1

### Purpose
Identifies locked user accounts across all domain controllers, determines lockout sources, and optionally unlocks accounts with comprehensive reporting.

### Key Features
✅ Multi-DC scanning for locked accounts  
✅ Identifies lockout source computer  
✅ Shows bad password count and lockout time  
✅ Optional automatic unlock  
✅ Security team email alerts  
✅ Professional HTML report generation  
✅ Event log analysis for forensics  

### Prerequisites
- Active Directory PowerShell Module
- Domain Admin or Account Operator privileges
- Access to read Security event logs on DCs

### Basic Usage

```powershell
# Simple scan for locked accounts
.\Find-LockedAccounts.ps1 -Verbose

# Scan and unlock all locked accounts
.\Find-LockedAccounts.ps1 -UnlockAccounts

# Scan with email alert to security team
.\Find-LockedAccounts.ps1 -EmailAlert -EmailTo "security@contoso.com" -SMTPServer "smtp.contoso.com"

# Full scan with unlock and reporting
.\Find-LockedAccounts.ps1 -UnlockAccounts -OutputReport -EmailAlert -EmailTo "security@contoso.com"
```

### Advanced Examples

```powershell
# Check specific domain
.\Find-LockedAccounts.ps1 -Domain "contoso.com" -OutputReport

# Generate report only (no unlock)
.\Find-LockedAccounts.ps1 -OutputReport -ReportPath "C:\Reports\LockedAccounts.html"

# Multiple email recipients
.\Find-LockedAccounts.ps1 -EmailAlert -EmailTo "security@contoso.com","admin@contoso.com"

# Unlock with confirmation prompt
.\Find-LockedAccounts.ps1 -UnlockAccounts -WhatIf  # Preview mode
.\Find-LockedAccounts.ps1 -UnlockAccounts -Confirm  # Interactive confirmation
```

### Output Files

| File | Location | Description |
|------|----------|-------------|
| Log File | `C:\Logs\LockedAccounts_YYYYMMDD_HHMMSS.log` | Detailed operation log |
| Error Log | `C:\Logs\Errors.txt` | Error details and stack traces |
| HTML Report | `C:\Logs\LockedAccounts_Report_YYYYMMDD.html` | Visual report with statistics |

### What It Does

1. **Discovery Phase**
   - Queries all domain controllers
   - Identifies locked accounts
   - Eliminates duplicates across DCs

2. **Analysis Phase**
   - Determines lockout source computer
   - Retrieves bad password count
   - Identifies originating DC
   - Checks Security event logs (Event ID 4740)

3. **Remediation Phase** (if `-UnlockAccounts`)
   - Unlocks each account
   - Logs all actions
   - Marks status in report

4. **Reporting Phase**
   - Generates HTML report
   - Sends email alerts (if configured)
   - Provides summary statistics

### Sample Output

```
[2025-01-15 09:30:15] [INFO] Starting Locked Account Detection
[2025-01-15 09:30:16] [SUCCESS] Found 3 domain controller(s)
[2025-01-15 09:30:20] [WARNING] Found 2 locked account(s)
[2025-01-15 09:30:20] [WARNING]   🔒 jdoe - Locked since 2025-01-15 08:45:00
[2025-01-15 09:30:20] [INFO]      Source: WORKSTATION-123
[2025-01-15 09:30:21] [SUCCESS] ✓ Unlocked account: jdoe
[2025-01-15 09:30:25] [SUCCESS] HTML report saved
[2025-01-15 09:30:26] [SUCCESS] Security alert email sent
```

### Security Considerations

- ⚠️ **Lockout sources** may indicate password spray attacks
- ⚠️ **Multiple consecutive lockouts** from same source = investigate
- ⚠️ **Service accounts** repeatedly locking = check application credentials
- ✅ **Always review** lockout patterns before mass unlocking

### Scheduled Task Setup

Run hourly to catch lockouts quickly:

```powershell
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
    -Argument '-ExecutionPolicy Bypass -File "C:\Scripts\Find-LockedAccounts.ps1" -EmailAlert -EmailTo "security@contoso.com"'

$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 1)

$Principal = New-ScheduledTaskPrincipal -UserId "DOMAIN\ServiceAccount" -LogonType Password -RunLevel Highest

Register-ScheduledTask -TaskName "Monitor Locked Accounts" `
    -Action $Action -Trigger $Trigger -Principal $Principal
```

---

## 📧 Send-PasswordExpirationNotification.ps1

### Purpose
Proactive password expiration notification system that sends branded emails to users before their passwords expire, reducing helpdesk calls and lockouts.

### Key Features
✅ Configurable warning thresholds (default: 14 days)  
✅ Beautiful, branded HTML email notifications  
✅ Manager CC option  
✅ Urgency-based color coding  
✅ Test mode for validation  
✅ Excludes disabled users and service accounts  
✅ Comprehensive HTML reporting  
✅ Step-by-step password change instructions  

### Prerequisites
- Active Directory PowerShell Module
- SMTP server access
- Users must have email addresses in AD

### Basic Usage

```powershell
# Send notifications for passwords expiring in 14 days
.\Send-PasswordExpirationNotification.ps1

# Custom threshold (7 days)
.\Send-PasswordExpirationNotification.ps1 -DaysBeforeExpiration 7

# Include manager on notifications
.\Send-PasswordExpirationNotification.ps1 -NotifyManager

# Test mode (no emails sent)
.\Send-PasswordExpirationNotification.ps1 -TestMode -GenerateReport
```

### Advanced Examples

```powershell
# Full production run
.\Send-PasswordExpirationNotification.ps1 -DaysBeforeExpiration 14 `
    -NotifyManager `
    -SMTPServer "smtp.contoso.com" `
    -EmailFrom "noreply@contoso.com" `
    -GenerateReport

# Exclude service account OUs
.\Send-PasswordExpirationNotification.ps1 -ExcludeOUs @(
    "OU=Service Accounts,DC=contoso,DC=com",
    "OU=Admin Accounts,DC=contoso,DC=com"
)

# Custom report location
.\Send-PasswordExpirationNotification.ps1 -GenerateReport `
    -ReportPath "C:\Reports\PasswordExpiration_$(Get-Date -Format 'yyyyMMdd').html"

# Multiple warning intervals
# Run daily with different thresholds
.\Send-PasswordExpirationNotification.ps1 -DaysBeforeExpiration 3  # Urgent
.\Send-PasswordExpirationNotification.ps1 -DaysBeforeExpiration 7  # Important
.\Send-PasswordExpirationNotification.ps1 -DaysBeforeExpiration 14 # Reminder
```

### Notification Urgency Levels

| Days Until Expiry | Priority | Email Color | Subject |
|-------------------|----------|-------------|---------|
| 1-3 days | 🔴 URGENT | Red | [URGENT] Your password expires in X days |
| 4-7 days | ⚠️ IMPORTANT | Yellow | [IMPORTANT] Your password expires in X days |
| 8+ days | ℹ️ REMINDER | Blue | [REMINDER] Your password expires in X days |

### Email Features

**User-Friendly Content:**
- Clear expiration countdown
- Multiple password change methods (Ctrl+Alt+Del, Online portal)
- Password requirements and tips
- Strong password examples
- IT support contact information

**Visual Design:**
- Professional gradient header
- Color-coded urgency banners
- Responsive HTML layout
- Company branding ready

### Output Files

| File | Location | Description |
|------|----------|-------------|
| Log File | `C:\Logs\PasswordExpiration_YYYYMMDD_HHMMSS.log` | Detailed execution log |
| Error Log | `C:\Logs\Errors.txt` | Failed notifications |
| HTML Report | `C:\Logs\PasswordExpiration_Report_YYYYMMDD.html` | Summary report |

### What It Does

1. **Policy Check**
   - Retrieves domain password policy
   - Calculates maximum password age
   - Validates expiration is enabled

2. **User Discovery**
   - Gets all enabled users
   - Filters users with non-expiring passwords
   - Excludes specified OUs
   - Validates email addresses

3. **Expiration Calculation**
   - Calculates days until expiration
   - Filters by threshold
   - Categorizes by urgency

4. **Notification Delivery**
   - Sends personalized emails
   - Includes manager (if configured)
   - Tracks success/failure
   - Rate-limits to avoid SMTP overload

5. **Reporting**
   - Generates HTML summary
   - Statistics by urgency level
   - Success/failure tracking

### Sample Output

```
[2025-01-15 08:00:00] [INFO] Password Expiration Notification System
[2025-01-15 08:00:01] [INFO] Maximum password age: 90 days
[2025-01-15 08:00:02] [SUCCESS] Found 150 enabled users
[2025-01-15 08:00:05] [WARNING] Found 12 user(s) with passwords expiring
[2025-01-15 08:00:06] [SUCCESS] Email sent to: john.doe@contoso.com
[2025-01-15 08:00:06] [SUCCESS] Email sent to: jane.smith@contoso.com
[2025-01-15 08:00:10] [SUCCESS] Notifications sent: 12 / 12
[2025-01-15 08:00:11] [INFO] Breakdown by urgency:
[2025-01-15 08:00:11] [ERROR]   🔴 Urgent (≤3 days): 2
[2025-01-15 08:00:11] [WARNING]   ⚠️  Important (4-7 days): 5
[2025-01-15 08:00:11] [INFO]   ℹ️  Reminder (8+ days): 5
```

### Scheduled Task Setup

Run daily at 8:00 AM:

```powershell
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
    -Argument '-ExecutionPolicy Bypass -File "C:\Scripts\Send-PasswordExpirationNotification.ps1" -DaysBeforeExpiration 14 -NotifyManager -GenerateReport'

$Trigger = New-ScheduledTaskTrigger -Daily -At "08:00AM"

$Principal = New-ScheduledTaskPrincipal -UserId "DOMAIN\ServiceAccount" -LogonType Password

Register-ScheduledTask -TaskName "Password Expiration Notifications" `
    -Action $Action -Trigger $Trigger -Principal $Principal `
    -Description "Daily password expiration notifications"
```

### Best Practices

✅ **Test First**: Always run with `-TestMode` initially  
✅ **Multiple Intervals**: Run at 3, 7, and 14 days for best coverage  
✅ **Exclude Service Accounts**: Use `-ExcludeOUs` for system accounts  
✅ **Monitor Delivery**: Check error logs for failed notifications  
✅ **Update Branding**: Customize HTML templates with company branding  
✅ **Track Metrics**: Use reports to measure effectiveness  

### Troubleshooting

**No emails sent:**
- Verify SMTP server is reachable
- Check service account has send permissions
- Validate user email addresses in AD

**Users not receiving emails:**
- Check spam/junk folders
- Verify email address format
- Check SMTP relay logs

**High failure rate:**
- Check error log: `C:\Logs\Errors.txt`
- Verify email addresses in AD
- Test SMTP connectivity

---

## 🔄 Integration Example

Use both scripts together for comprehensive account security:

```powershell
# Daily at 8:00 AM - Password notifications
.\Send-PasswordExpirationNotification.ps1 -DaysBeforeExpiration 14 `
    -NotifyManager -GenerateReport

# Hourly - Check for locked accounts
.\Find-LockedAccounts.ps1 -EmailAlert -EmailTo "security@contoso.com" `
    -OutputReport
```

---

## 📊 Expected Results

### Find-LockedAccounts.ps1
- **Reduces mean time to unlock**: From hours to minutes
- **Identifies security threats**: Password spray attacks, compromised credentials
- **Improves user experience**: Faster resolution of lockouts
- **Audit compliance**: Complete lockout history and remediation

### Send-PasswordExpirationNotification.ps1
- **Reduces helpdesk calls**: 50-70% reduction in password-related tickets
- **Prevents lockouts**: Users change passwords before expiry
- **Improves security**: Users prepared to create strong passwords
- **Better user experience**: Proactive communication

---

## 🎓 Tips for Success

1. **Start with Test Mode**: Validate email formatting and recipients
2. **Customize Templates**: Add company logo and branding
3. **Monitor Logs**: Review daily for trends and issues
4. **Adjust Thresholds**: Fine-tune based on your environment
5. **Communicate**: Let users know about the new system
6. **Track Metrics**: Measure reduction in helpdesk calls
7. **Schedule Wisely**: Run notifications early morning, lockout checks hourly

---

## 📞 Support

For issues or questions:
- Review log files in `C:\Logs\`
- Check error logs for detailed error messages
- Verify prerequisites are met
- Test connectivity to SMTP/DCs

---

**Last Updated**: January 2025  
**Scripts Version**: 1.0  
**Compatibility**: Windows Server 2016+, PowerShell 5.1+