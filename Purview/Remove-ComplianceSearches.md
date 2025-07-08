# Remove Compliance Searches PowerShell Script

## Overview

This PowerShell script automates the removal of Microsoft 365 compliance searches from the Security & Compliance Center using a CSV file input. It provides safe deletion capabilities with comprehensive error handling, logging, and validation features.

## Features

- **Bulk Removal**: Process multiple compliance searches from a CSV file
- **Safety Mode**: Preview operations with `-WhatIf` parameter before execution
- **Comprehensive Validation**: Checks for search existence, running status, and proper permissions
- **Detailed Logging**: Color-coded output with timestamp logging for audit trails
- **Error Handling**: Graceful handling of errors with detailed error messages
- **Connection Verification**: Automatically validates Security & Compliance Center connection
- **Flexible CSV Format**: Supports both "SearchName" and "Name" column headers

## Prerequisites

### PowerShell Modules
- **Exchange Online PowerShell V2** or **Microsoft 365 Security & Compliance PowerShell**
- PowerShell 5.1 or later (PowerShell 7+ recommended)

### Installation
```powershell
# Install required module
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser

# Alternative: Install Security & Compliance module
Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```

### Authentication & Permissions
- **Required Role**: eDiscovery Manager or Compliance Administrator
- **Connection**: Must be connected to Security & Compliance Center PowerShell

```powershell
# Connect to Security & Compliance Center
Connect-IPPSSession -UserPrincipalName admin@contoso.com
```

## CSV File Format

The script accepts CSV files with either `SearchName` or `Name` as the column header:

### Example CSV Format:
```csv
SearchName
"HR Investigation 2024-Q1"
"Legal Hold - Employee Departure"
"Compliance Audit - Finance Department"
"eDiscovery Case 12345"
```

### Alternative Format:
```csv
Name
"Security Incident Response"
"Regulatory Compliance Check"
"Internal Audit Search"
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `CsvFilePath` | String | Yes | Full path to the CSV file containing search names |
| `WhatIf` | Switch | No | Preview mode - shows what would be deleted without executing |
| `Force` | Switch | No | Skip confirmation prompts during deletion |

## Usage Examples

### 1. Preview Mode (Recommended First Run)
```powershell
.\Remove-ComplianceSearches.ps1 -CsvFilePath "C:\ComplianceSearches.csv" -WhatIf
```

### 2. Interactive Deletion with Confirmations
```powershell
.\Remove-ComplianceSearches.ps1 -CsvFilePath "C:\ComplianceSearches.csv"
```

### 3. Automated Deletion (No Confirmations)
```powershell
.\Remove-ComplianceSearches.ps1 -CsvFilePath "C:\ComplianceSearches.csv" -Force
```

### 4. Using Relative Paths
```powershell
.\Remove-ComplianceSearches.ps1 -CsvFilePath ".\searches\cleanup_list.csv" -WhatIf
```

## Script Behavior

### Validation Checks
1. **CSV File Existence**: Verifies the CSV file exists at the specified path
2. **Column Validation**: Confirms CSV contains "SearchName" or "Name" column
3. **Connection Status**: Tests Security & Compliance Center PowerShell connection
4. **Search Existence**: Verifies each compliance search exists before attempting deletion
5. **Search Status**: Checks if searches are currently running (cannot delete active searches)

### Processing Logic
- **Empty Names**: Skips rows with empty or whitespace-only search names
- **Missing Searches**: Logs warning for searches that don't exist
- **Running Searches**: Skips searches with "InProgress" status
- **Error Handling**: Continues processing remaining searches if individual failures occur

## Output and Logging

### Console Output
The script provides color-coded console output:
- **Green**: Successful operations
- **Red**: Errors and failures
- **Yellow**: Warnings and skipped items
- **Cyan**: What-if mode operations
- **White**: General information

### Log Format
```
[2024-12-15 14:30:15] [INFO] Reading CSV file: C:\searches.csv
[2024-12-15 14:30:16] [INFO] Found 5 compliance searches to process
[2024-12-15 14:30:17] [INFO] Processing: HR Investigation 2024-Q1
```

### Summary Report
At completion, the script displays:
- Total searches processed
- Successful deletions
- Failed operations
- Skipped items (with reasons)

## Error Handling

### Common Errors and Solutions

| Error | Cause | Solution |
|-------|-------|----------|
| "CSV file not found" | Invalid file path | Verify file path and permissions |
| "Not connected to Security & Compliance Center" | No PowerShell session | Run `Connect-IPPSSession` |
| "Search not found" | Search name doesn't exist | Verify search names in CSV |
| "Search is currently running" | Search in progress | Wait for completion or stop search |
| "Access denied" | Insufficient permissions | Ensure eDiscovery Manager role |

### Exit Codes
- **0**: Success
- **1**: Critical error (file not found, connection issues, CSV format errors)

## Security Considerations

### Permissions
- Script requires **eDiscovery Manager** or **Compliance Administrator** role
- Follows principle of least privilege - only removes searches, cannot create or modify

### Audit Trail
- All operations are logged with timestamps
- Failed operations are recorded with error details
- What-if mode provides safe preview without changes

### Data Protection
- No sensitive data is stored in logs
- CSV file should be secured with appropriate file permissions
- Consider using encrypted storage for CSV files containing sensitive search names

## Best Practices

### Before Running
1. **Always test with `-WhatIf`** parameter first
2. **Backup** important searches before bulk deletion
3. **Verify CSV format** and search names
4. **Confirm permissions** and connection status

### During Execution
1. **Monitor output** for errors or warnings
2. **Keep CSV file** for audit purposes
3. **Document deletions** for compliance requirements

### After Completion
1. **Review summary report** for any failures
2. **Verify deletions** in Security & Compliance Center
3. **Update documentation** with completed actions

## Troubleshooting

### Connection Issues
```powershell
# Test connection
Get-ComplianceSearch -ResultSize 1

# Reconnect if needed
Disconnect-ExchangeOnline
Connect-IPPSSession
```

### CSV Format Issues
```powershell
# Check CSV structure
Import-Csv "C:\searches.csv" | Select-Object -First 5
```

### Permission Verification
```powershell
# Check current user roles
Get-RoleGroupMember "eDiscovery Manager"
```

## Advanced Usage

### Custom Log File
Uncomment the log file section in the script to enable file logging:
```powershell
# $logMessage | Out-File -FilePath "compliance_search_removal.log" -Append
```

### Filtering by Status
Pre-filter searches before creating CSV:
```powershell
Get-ComplianceSearch | Where-Object {$_.Status -eq "Completed"} | 
    Select-Object Name | Export-Csv -Path "completed_searches.csv" -NoTypeInformation
```

## Version History

- **v1.0**: Initial release with basic functionality
- **v1.1**: Added What-if mode and improved error handling
- **v1.2**: Enhanced logging and validation features

## Support

For issues related to:
- **Script functionality**: Review error messages and troubleshooting section
- **Microsoft 365 permissions**: Contact your Microsoft 365 administrator
- **Compliance requirements**: Consult your organization's compliance team

## License

This script is provided as-is for educational and administrative purposes. Test thoroughly in your environment before production use.
