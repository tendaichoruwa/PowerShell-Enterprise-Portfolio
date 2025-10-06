<#
.SYNOPSIS
Create the repository directory structure.

.DESCRIPTION
Idempotently creates a set of directories under the specified root (defaults to the script folder).

.EXAMPLE
.\directories.ps1
.EXAMPLE
.\directories.ps1 -Root "C:\Repos\MyPortfolio"
#>

[CmdletBinding()]
param(
    # Base folder to create the structure in
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $Root = $(if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path })
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Define the directory structure (relative to $Root)
$Directories = @(
    'ActiveDirectory/UserManagement',
    'ActiveDirectory/HealthMonitoring',
    'ActiveDirectory/GroupManagement',
    'EntraID-M365/UserManagement',
    'EntraID-M365/Compliance',
    'Intune/DeviceManagement',
    'Intune/Compliance',
    'HyperV/VMManagement',
    'Common/Modules',
    'Common/Templates',
    'Examples/SampleData',
    'Examples/Workflows',
    'Tests/Pester',
    'Docs',
    '.github/workflows'
)

Write-Host "Creating directory structure under: $Root" -ForegroundColor Cyan

$created = 0
$existing = 0

foreach ($rel in $Directories) {
    # Normalize to full path under $Root
    $full = Join-Path -Path $Root -ChildPath $rel

    if (-not (Test-Path -LiteralPath $full)) {
        New-Item -ItemType Directory -Path $full -Force | Out-Null
        Write-Host "[+] Created: $rel" -ForegroundColor Green
        $created++
    }
    else {
        Write-Host "[=] Exists : $rel" -ForegroundColor DarkYellow
        $existing++
    }
}

Write-Host ""
Write-Host "Done. Created: $created, Existing: $existing, Total: $($Directories.Count)" -ForegroundColor Green
