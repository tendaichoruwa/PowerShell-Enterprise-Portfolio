<#
.SYNOPSIS
    Comprehensive Domain Controller replication health monitoring script.

.DESCRIPTION
    This script performs thorough Active Directory replication health checks including:
    - Querying all domain controllers
    - Checking replication status between all DCs
    - Identifying replication failures and latency issues
    - Generating HTML status report
    - Sending email alerts for critical issues

.PARAMETER Domain
    Fully qualified domain name. If not specified, uses current domain.

.PARAMETER EmailReport
    If specified, sends the HTML report via email.

.PARAMETER EmailTo
    Email address(es) to send the report to (comma-separated).

.PARAMETER SMTPServer
    SMTP server address for sending reports.

.PARAMETER AlertThresholdMinutes
    Replication latency threshold in minutes for alerts. Default: 60 minutes.

.PARAMETER OutputPath
    Path to save the HTML report. Default: C:\Logs\DCReplication_Report_YYYYMMDD.html

.EXAMPLE
    .\Get-DCReplicationStatus.ps1 -Verbose

.EXAMPLE
    .\Get-DCReplicationStatus.ps1 -EmailReport -EmailTo "admin@contoso.com" -SMTPServer "smtp.contoso.com"

.EXAMPLE
    .\Get-DCReplicationStatus.ps1 -AlertThresholdMinutes 30 -OutputPath "C:\Reports\DCHealth.html"

.NOTES
    Author: Tendai Choruwa
    Version: 1.0
    Last Updated: January 2025
    Requires: Active Directory PowerShell Module, Domain Admin or Enterprise Admin privileges
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$Domain,

    [Parameter(Mandatory = $false)]
    [switch]$EmailReport,

    [Parameter(Mandatory = $false)]
    [string[]]$EmailTo,

    [Parameter(Mandatory = $false)]
    [string]$SMTPServer = "smtp.company.com",

    [Parameter(Mandatory = $false)]
    [int]$AlertThresholdMinutes = 60,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
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

$LogFile = Join-Path -Path $LogDirectory -ChildPath "DCReplicationCheck_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$ErrorLogFile = Join-Path -Path $LogDirectory -ChildPath "Errors.txt"

if (-not $OutputPath) {
    $OutputPath = Join-Path -Path $LogDirectory -ChildPath "DCReplication_Report_$(Get-Date -Format 'yyyyMMdd').html"
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

function Get-ReplicationPartnerMetadata {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourceDC,
        
        [Parameter(Mandatory = $true)]
        [string]$DestinationDC
    )

    try {
        $repadmin = repadmin /showrepl $DestinationDC /verbose
        
        $metadata = [PSCustomObject]@{
            SourceDC            = $SourceDC
            DestinationDC       = $DestinationDC
            LastSuccessfulSync  = $null
            LastAttempt         = $null
            ConsecutiveFailures = 0
            LastFailureStatus   = "N/A"
            Status              = "Unknown"
        }

        # Parse repadmin output
        $inReplicationSection = $false
        foreach ($line in $repadmin) {
            if ($line -match $SourceDC) {
                $inReplicationSection = $true
            }
            
            if ($inReplicationSection) {
                if ($line -match "Last attempt @ (\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})") {
                    $metadata.LastAttempt = [datetime]::Parse($Matches[1])
                }
                if ($line -match "Last success @ (\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})") {
                    $metadata.LastSuccessfulSync = [datetime]::Parse($Matches[1])
                }
                if ($line -match "consecutive failure\(s\)\.(\d+)") {
                    $metadata.ConsecutiveFailures = [int]$Matches[1]
                }
                if ($line -match "Last error: (\d+)") {
                    $metadata.LastFailureStatus = $Matches[1]
                }
            }
        }

        # Determine status
        if ($metadata.ConsecutiveFailures -gt 0) {
            $metadata.Status = "Failed"
        }
        elseif ($metadata.LastSuccessfulSync) {
            $timeSinceSync = (Get-Date) - $metadata.LastSuccessfulSync
            if ($timeSinceSync.TotalMinutes -gt $AlertThresholdMinutes) {
                $metadata.Status = "Warning"
            }
            else {
                $metadata.Status = "Healthy"
            }
        }

        return $metadata
    }
    catch {
        Write-Log -Message "Error getting replication metadata for $SourceDC -> $DestinationDC : $($_.Exception.Message)" -Level "WARNING"
        return $null
    }
}

function New-HTMLReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array]$DomainControllers,
        
        [Parameter(Mandatory = $true)]
        [array]$ReplicationData,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Summary
    )

    $healthyCount = ($ReplicationData | Where-Object { $_.Status -eq "Healthy" }).Count
    $warningCount = ($ReplicationData | Where-Object { $_.Status -eq "Warning" }).Count
    $failedCount = ($ReplicationData | Where-Object { $_.Status -eq "Failed" }).Count
    
    $statusColor = if ($failedCount -gt 0) { "#dc3545" } elseif ($warningCount -gt 0) { "#ffc107" } else { "#28a745" }
    $statusText = if ($failedCount -gt 0) { "CRITICAL" } elseif ($warningCount -gt 0) { "WARNING" } else { "HEALTHY" }

    $dcRows = ""
    foreach ($dc in $DomainControllers) {
        $pingStatus = if (Test-Connection -ComputerName $dc.HostName -Count 1 -Quiet) { "✓ Online" } else { "✗ Offline" }
        $pingColor = if ($pingStatus -match "Online") { "#28a745" } else { "#dc3545" }
        
        $dcRows += @"
        <tr>
            <td>$($dc.HostName)</td>
            <td>$($dc.Site)</td>
            <td>$($dc.OperatingSystem)</td>
            <td style="color: $pingColor; font-weight: bold;">$pingStatus</td>
        </tr>
"@
    }

    $replRows = ""
    foreach ($repl in $ReplicationData) {
        $rowColor = switch ($repl.Status) {
            "Healthy" { "#d4edda" }
            "Warning" { "#fff3cd" }
            "Failed"  { "#f8d7da" }
            default   { "#ffffff" }
        }
        
        $statusBadge = switch ($repl.Status) {
            "Healthy" { "<span style='background-color: #28a745; color: white; padding: 4px 8px; border-radius: 4px;'>✓ Healthy</span>" }
            "Warning" { "<span style='background-color: #ffc107; color: black; padding: 4px 8px; border-radius: 4px;'>⚠ Warning</span>" }
            "Failed"  { "<span style='background-color: #dc3545; color: white; padding: 4px 8px; border-radius: 4px;'>✗ Failed</span>" }
            default   { "<span style='background-color: #6c757d; color: white; padding: 4px 8px; border-radius: 4px;'>? Unknown</span>" }
        }
        
        $lastSync = if ($repl.LastSuccessfulSync) { $repl.LastSuccessfulSync.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
        $timeSinceSync = if ($repl.LastSuccessfulSync) { 
            $span = (Get-Date) - $repl.LastSuccessfulSync
            "$([math]::Round($span.TotalMinutes, 0)) minutes ago"
        } else { 
            "N/A" 
        }
        
        $replRows += @"
        <tr style="background-color: $rowColor;">
            <td>$($repl.SourceDC)</td>
            <td>$($repl.DestinationDC)</td>
            <td>$lastSync</td>
            <td>$timeSinceSync</td>
            <td>$($repl.ConsecutiveFailures)</td>
            <td style="text-align: center;">$statusBadge</td>
        </tr>
"@
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Domain Controller Replication Health Report</title>
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
        .status-banner {
            background-color: $statusColor;
            color: white;
            padding: 20px;
            text-align: center;
            font-size: 24px;
            font-weight: bold;
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
        .alert {
            background-color: #fff3cd;
            border-left: 4px solid #ffc107;
            padding: 15px;
            margin: 20px 0;
            border-radius: 4px;
        }
        .alert-danger {
            background-color: #f8d7da;
            border-left: 4px solid #dc3545;
        }
        .alert-success {
            background-color: #d4edda;
            border-left: 4px solid #28a745;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🖥️ Domain Controller Replication Health Report</h1>
            <p>Generated on $(Get-Date -Format "MMMM dd, yyyy 'at' HH:mm:ss")</p>
        </div>
        
        <div class="status-banner">
            OVERALL STATUS: $statusText
        </div>
        
        <div class="summary">
            <div class="summary-card">
                <h3>Total Domain Controllers</h3>
                <div class="value">$($Summary.TotalDCs)</div>
            </div>
            <div class="summary-card">
                <h3>Healthy Connections</h3>
                <div class="value" style="color: #28a745;">$healthyCount</div>
            </div>
            <div class="summary-card">
                <h3>Warnings</h3>
                <div class="value" style="color: #ffc107;">$warningCount</div>
            </div>
            <div class="summary-card">
                <h3>Failed Replications</h3>
                <div class="value" style="color: #dc3545;">$failedCount</div>
            </div>
            <div class="summary-card">
                <h3>Total Repl. Partners</h3>
                <div class="value">$($Summary.TotalPartners)</div>
            </div>
        </div>
        
        $(if ($failedCount -gt 0) {
            "<div class='section'><div class='alert alert-danger'><strong>⚠️ CRITICAL ALERT:</strong> $failedCount replication connection(s) have failed. Immediate attention required!</div></div>"
        } elseif ($warningCount -gt 0) {
            "<div class='section'><div class='alert'><strong>⚠️ WARNING:</strong> $warningCount replication connection(s) exceed the latency threshold of $AlertThresholdMinutes minutes.</div></div>"
        } else {
            "<div class='section'><div class='alert alert-success'><strong>✓ ALL SYSTEMS HEALTHY:</strong> All domain controller replication connections are functioning normally.</div></div>"
        })
        
        <div class="section">
            <h2>📊 Domain Controllers Inventory</h2>
            <table>
                <thead>
                    <tr>
                        <th>Hostname</th>
                        <th>AD Site</th>
                        <th>Operating System</th>
                        <th>Network Status</th>
                    </tr>
                </thead>
                <tbody>
                    $dcRows
                </tbody>
            </table>
        </div>
        
        <div class="section">
            <h2>🔄 Replication Status Details</h2>
            <table>
                <thead>
                    <tr>
                        <th>Source DC</th>
                        <th>Destination DC</th>
                        <th>Last Successful Sync</th>
                        <th>Time Since Sync</th>
                        <th>Consecutive Failures</th>
                        <th style="text-align: center;">Status</th>
                    </tr>
                </thead>
                <tbody>
                    $replRows
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <p><strong>AD Replication Health Monitor</strong></p>
            <p>Report generated by PowerShell automation script</p>
            <p>Log file: $LogFile</p>
            <p>© $(Get-Date -Format yyyy) IT Operations | Threshold: $AlertThresholdMinutes minutes</p>
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
Write-Log -Message "Starting DC Replication Health Check" -Level "INFO"
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
    $DomainControllers = Get-ADDomainController -Filter * -Server $Domain | Select-Object HostName, Site, OperatingSystem, IPv4Address
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
# STEP 2: CHECK REPLICATION STATUS
# ============================================================================

Write-Log -Message "Step 2: Checking replication status..." -Level "INFO"

$ReplicationData = @()
$TotalPartners = 0

foreach ($destDC in $DomainControllers) {
    Write-Log -Message "Querying replication partners for $($destDC.HostName)..." -Level "INFO"
    
    try {
        # Get replication partners
        $partners = Get-ADReplicationPartnerMetadata -Target $destDC.HostName -ErrorAction Stop
        
        foreach ($partner in $partners) {
            $TotalPartners++
            
            $timeSinceSync = if ($partner.LastReplicationSuccess) {
                ((Get-Date) - $partner.LastReplicationSuccess).TotalMinutes
            } else {
                999999
            }
            
            $status = if ($partner.ConsecutiveReplicationFailures -gt 0) {
                "Failed"
            } elseif ($timeSinceSync -gt $AlertThresholdMinutes) {
                "Warning"
            } else {
                "Healthy"
            }
            
            $replData = [PSCustomObject]@{
                SourceDC            = $partner.Partner
                DestinationDC       = $destDC.HostName
                LastSuccessfulSync  = $partner.LastReplicationSuccess
                LastAttempt         = $partner.LastReplicationAttempt
                ConsecutiveFailures = $partner.ConsecutiveReplicationFailures
                LastFailureStatus   = $partner.LastReplicationResult
                Status              = $status
            }
            
            $ReplicationData += $replData
            
            $logLevel = switch ($status) {
                "Failed"  { "ERROR" }
                "Warning" { "WARNING" }
                default   { "SUCCESS" }
            }
            
            Write-Log -Message "  $($partner.Partner) -> $($destDC.HostName): $status" -Level $logLevel
        }
    }
    catch {
        $ErrorMessage = "Failed to get replication data for $($destDC.HostName): $($_.Exception.Message)"
        Write-Log -Message $ErrorMessage -Level "WARNING"
        Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    }
}

# ============================================================================
# STEP 3: GENERATE SUMMARY STATISTICS
# ============================================================================

Write-Log -Message "Step 3: Generating summary statistics..." -Level "INFO"

$Summary = @{
    TotalDCs         = $DomainControllers.Count
    TotalPartners    = $TotalPartners
    HealthyCount     = ($ReplicationData | Where-Object { $_.Status -eq "Healthy" }).Count
    WarningCount     = ($ReplicationData | Where-Object { $_.Status -eq "Warning" }).Count
    FailedCount      = ($ReplicationData | Where-Object { $_.Status -eq "Failed" }).Count
    CheckTime        = Get-Date
}

Write-Log -Message "Summary Statistics:" -Level "INFO"
Write-Log -Message "  Total DCs: $($Summary.TotalDCs)" -Level "INFO"
Write-Log -Message "  Total Replication Partners: $($Summary.TotalPartners)" -Level "INFO"
Write-Log -Message "  Healthy: $($Summary.HealthyCount)" -Level "SUCCESS"
Write-Log -Message "  Warnings: $($Summary.WarningCount)" -Level $(if ($Summary.WarningCount -gt 0) { "WARNING" } else { "INFO" })
Write-Log -Message "  Failed: $($Summary.FailedCount)" -Level $(if ($Summary.FailedCount -gt 0) { "ERROR" } else { "INFO" })

# ============================================================================
# STEP 4: GENERATE HTML REPORT
# ============================================================================

Write-Log -Message "Step 4: Generating HTML report..." -Level "INFO"

try {
    $htmlReport = New-HTMLReport -DomainControllers $DomainControllers -ReplicationData $ReplicationData -Summary $Summary
    $htmlReport | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
    Write-Log -Message "HTML report saved to: $OutputPath" -Level "SUCCESS"
}
catch {
    $ErrorMessage = "Failed to generate HTML report: $($_.Exception.Message)"
    Write-Log -Message $ErrorMessage -Level "ERROR"
    Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
}

# ============================================================================
# STEP 5: SEND EMAIL REPORT (IF REQUESTED)
# ============================================================================

if ($EmailReport -and $EmailTo) {
    Write-Log -Message "Step 5: Sending email report..." -Level "INFO"
    
    try {
        $subject = "Domain Controller Replication Health Report - $(Get-Date -Format 'yyyy-MM-dd')"
        
        if ($Summary.FailedCount -gt 0) {
            $subject = "🔴 CRITICAL: $subject"
        } elseif ($Summary.WarningCount -gt 0) {
            $subject = "⚠️ WARNING: $subject"
        } else {
            $subject = "✅ HEALTHY: $subject"
        }
        
        $EmailParams = @{
            To         = $EmailTo
            From       = "dchealth@company.com"
            Subject    = $subject
            Body       = $htmlReport
            BodyAsHtml = $true
            SmtpServer = $SMTPServer
            Priority   = if ($Summary.FailedCount -gt 0) { "High" } else { "Normal" }
        }
        
        Send-MailMessage @EmailParams -ErrorAction Stop
        Write-Log -Message "Email report sent successfully to: $($EmailTo -join ', ')" -Level "SUCCESS"
    }
    catch {
        $ErrorMessage = "Failed to send email report: $($_.Exception.Message)"
        Write-Log -Message $ErrorMessage -Level "WARNING"
        Add-Content -Path $ErrorLogFile -Value "$(Get-Date) - $ErrorMessage"
    }
}

# ============================================================================
# FINAL SUMMARY
# ============================================================================

Write-Log -Message "========================================" -Level "INFO"
Write-Log -Message "DC Replication Health Check Complete" -Level "INFO"
Write-Log -Message "========================================" -Level "INFO"

$overallStatus = if ($Summary.FailedCount -gt 0) {
    Write-Log -Message "OVERALL STATUS: CRITICAL - Immediate attention required" -Level "ERROR"
    "CRITICAL"
} elseif ($Summary.WarningCount -gt 0) {
    Write-Log -Message "OVERALL STATUS: WARNING - Review recommended" -Level "WARNING"
    "WARNING"
} else {
    Write-Log -Message "OVERALL STATUS: HEALTHY - All systems normal" -Level "SUCCESS"
    "HEALTHY"
}

Write-Log -Message "Report saved to: $OutputPath" -Level "INFO"
Write-Log -Message "Log file: $LogFile" -Level "INFO"

# Return summary object
[PSCustomObject]@{
    CheckDate        = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Domain           = $Domain
    TotalDCs         = $Summary.TotalDCs
    TotalPartners    = $Summary.TotalPartners
    HealthyCount     = $Summary.HealthyCount
    WarningCount     = $Summary.WarningCount
    FailedCount      = $Summary.FailedCount
    OverallStatus    = $overallStatus
    ReportPath       = $OutputPath
    LogFile          = $LogFile
}