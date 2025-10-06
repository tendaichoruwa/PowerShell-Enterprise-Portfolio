# Installation Guide

## Prerequisites
- PowerShell 5.1 or PowerShell 7.0+
- Windows 10/11 or Windows Server 2016+
- Required PowerShell modules (see below)

## Module Installation
```powershell
# Active Directory
Install-Module -Name ActiveDirectory -Force

# Microsoft Graph
Install-Module -Name Microsoft.Graph -Force

# Exchange Online
Install-Module -Name ExchangeOnlineManagement -Force
```

## Initial Setup
1. Clone the repository
2. Review scripts for your environment
3. Update variables (domain names, email servers, etc.)
4. Test in a non-production environment first