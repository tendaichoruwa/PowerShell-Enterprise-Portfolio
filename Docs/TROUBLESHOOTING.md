# Troubleshooting Guide

## Common Issues

### Issue: "Access Denied" errors
**Solution:** Ensure you have appropriate permissions (Domain Admin, Global Admin, etc.)

### Issue: Module not found
**Solution:** Install required modules using Install-Module

### Issue: Script execution disabled
**Solution:** Set execution policy: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

### Issue: Cannot connect to Microsoft Graph
**Solution:** Authenticate first: `Connect-MgGraph -Scopes "User.ReadWrite.All"`

## Getting Help
- Review script comments and help: `Get-Help .\ScriptName.ps1 -Full`
- Check log files in C:\Logs\
- Review error details in C:\Logs\Errors.txt