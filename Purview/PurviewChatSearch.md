# PurviewChatSearch_60DaysOld.ps1

## Overview

This PowerShell script creates a Microsoft Purview compliance search to find Microsoft Teams chats that are older than a specified number of days (default: 60 days). It processes a CSV file containing user information and creates a single bulk search across all specified users' mailboxes.

## Purpose

- **Compliance Management**: Identify old Teams chat data for retention policy enforcement
- **Data Governance**: Locate Teams chats that may need to be archived or purged
- **Legal Discovery**: Find historical Teams conversations for eDiscovery purposes
- **Storage Management**: Identify old chat data that may be consuming storage space

## Prerequisites

### Required PowerShell Modules
The script requires the following PowerShell modules to be installed:

```powershell
# Install required modules (run as Administrator)
Install-Module -Name ExchangeOnlineManagement -Force
Install-Module -Name Microsoft.Graph.Authentication -Force
Install-Module -Name Microsoft.Graph.Compliance -Force
```

### Required Permissions
The account running the script must have:
- **eDiscovery Manager** role in Microsoft 365 Security & Compliance Center
- **Global Administrator** or **Compliance Administrator** role (recommended)
- **Exchange Online Administrator** role (minimum)

### CSV File Format
The script expects a CSV file with user information. The CSV must contain one of the following column combinations:

**Option 1: Standard Format**
```csv
UserPrincipalName,DisplayName
john.doe@company.com,John Doe
jane.smith@company.com,Jane Smith
```

**Option 2: Using Email Column**
```csv
Email,DisplayName
john.doe@company.com,John Doe
jane.smith@company.com,Jane Smith
```

**Option 3: Using UPN Column**
```csv
UPN,Name
john.doe@company.com,John Doe
jane.smith@company.com,Jane Smith
```

## Parameters

### Required Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `CsvFilePath` | String | Path to the CSV file containing user information |

### Optional Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `TenantId` | String | None | Azure AD Tenant ID (required for app authentication) |
| `ClientId` | String | None | Azure AD Application Client ID (for app authentication) |
| `ClientSecret` | String | None | Azure AD Application Client Secret (for app authentication) |
| `SearchName` | String | Auto-generated | Custom name for the compliance search |
| `DaysOld` | Integer | 60 | Number of days old for chat filtering |

## Usage Examples

### Basic Usage (Interactive Authentication)
```powershell
# Search for Teams chats older than 60 days using interactive authentication
.\PurviewChatSearch_60DaysOld.ps1 -CsvFilePath "C:\Users\admin\users.csv"
```

### Custom Days Old
```powershell
# Search for Teams chats older than 90 days
.\PurviewChatSearch_60DaysOld.ps1 -CsvFilePath "C:\Users\admin\users.csv" -DaysOld 90

# Search for Teams chats older than 30 days
.\PurviewChatSearch_60DaysOld.ps1 -CsvFilePath "C:\Users\admin\users.csv" -DaysOld 30
```

### With Custom Search Name
```powershell
# Use a custom search name for better identification
.\PurviewChatSearch_60DaysOld.ps1 -CsvFilePath "C:\Users\admin\users.csv" -SearchName "Q4_2024_OldTeamsChats"
```

### Application Authentication (Service Principal)
```powershell
# Use service principal authentication for automated scenarios
.\PurviewChatSearch_60DaysOld.ps1 -CsvFilePath "C:\Users\admin\users.csv" `
    -TenantId "12345678-1234-1234-1234-123456789012" `
    -ClientId "87654321-4321-4321-4321-210987654321" `
    -ClientSecret "YourClientSecretHere"
```

### Complete Example with All Parameters
```powershell
# Full parameter usage
.\PurviewChatSearch_60DaysOld.ps1 `
    -CsvFilePath "C:\compliance\users_to_search.csv" `
    -TenantId "12345678-1234-1234-1234-123456789012" `
    -ClientId "87654321-4321-4321-4321-210987654321" `
    -ClientSecret "YourClientSecretHere" `
    -SearchName "Compliance_TeamsChats_OlderThan45Days" `
    -DaysOld 45
```

## Authentication Options

### Interactive Authentication (Default)
If no authentication parameters are provided, the script will prompt for interactive sign-in:
```powershell
.\PurviewChatSearch_60DaysOld.ps1 -CsvFilePath "users.csv"
```

### Service Principal Authentication
For automated scenarios, use application authentication:
```powershell
.\PurviewChatSearch_60DaysOld.ps1 -CsvFilePath "users.csv" `
    -TenantId "your-tenant-id" `
    -ClientId "your-client-id" `
    -ClientSecret "your-client-secret"
```

## How It Works

1. **CSV Processing**: Reads and validates the user list from the CSV file
2. **Authentication**: Connects to Microsoft Graph and Security & Compliance Center
3. **Date Calculation**: Calculates the cutoff date based on the `DaysOld` parameter
4. **Search Creation**: Creates a KQL query: `kind:microsoftteams AND sent<YYYY-MM-DD`
5. **Search Execution**: Starts the compliance search and monitors progress
6. **Results**: Displays search statistics and next steps

## KQL Query Details

The script generates a KQL (Keyword Query Language) query to find Teams chats:

```kql
kind:microsoftteams AND sent<2024-05-09
```

Where:
- `kind:microsoftteams` - Targets only Microsoft Teams messages
- `sent<YYYY-MM-DD` - Filters for messages sent before the cutoff date

## Fallback Strategies

The script includes multiple fallback strategies if the primary search fails:

1. **Primary**: `sent<YYYY-MM-DD` format
2. **Fallback 1**: `sent<MM/DD/YYYY` format
3. **Fallback 2**: Basic Teams search without date filter (with warning)

## Output and Results

### Console Output
The script provides detailed console output including:
- Connection status
- User validation results
- Search creation confirmation
- Real-time search progress
- Final results summary

### Search Results Summary
Upon completion, the script displays:
- Search name and status
- Total items found
- Total data size
- Number of users searched
- Search criteria used
- Cutoff date information

## Next Steps After Script Completion

1. **Access Compliance Center**:
   - Navigate to [Microsoft Purview Compliance Center](https://compliance.microsoft.com)
   - Go to "Content Search" under "Solutions"

2. **Review Results**:
   - Find your search by name
   - Review the search statistics
   - Examine the results preview

3. **Export Data** (if needed):
   - Select your search
   - Click "Export results"
   - Choose export options
   - Download the exported data

4. **Take Action**:
   - Create retention policies
   - Implement legal holds
   - Schedule data purging

## Error Handling

The script includes comprehensive error handling:
- **Module Loading**: Validates required PowerShell modules
- **CSV Validation**: Checks file existence and format
- **Authentication**: Handles connection failures
- **Search Creation**: Attempts multiple fallback strategies
- **Monitoring**: Tracks search progress with timeout handling

## Performance Considerations

- **Large User Lists**: The script can handle hundreds of users in a single search
- **Search Timeout**: Maximum wait time is 2 hours for search completion
- **Progress Monitoring**: Status updates every 2 minutes during search execution
- **Fallback Options**: Multiple strategies ensure search creation success

## Security Considerations

- **Credentials**: Use service principal authentication for automated scenarios
- **Permissions**: Follow principle of least privilege
- **Client Secrets**: Store secrets securely, avoid hardcoding
- **Audit Trail**: All actions are logged in Microsoft 365 audit logs

## Troubleshooting

### Common Issues and Solutions

**Module Import Errors**:
```powershell
# Run as Administrator
Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
```

**Authentication Failures**:
- Verify account has required permissions
- Check if MFA is properly configured
- Ensure service principal has correct API permissions

**CSV Format Issues**:
- Verify column names match expected format
- Check for empty rows or invalid email addresses
- Ensure UTF-8 encoding without BOM

**Search Creation Failures**:
- The script includes automatic fallback strategies
- Check compliance center permissions
- Verify tenant has appropriate licenses

### Debug Mode
For additional troubleshooting information, run with PowerShell verbose output:
```powershell
.\PurviewChatSearch_60DaysOld.ps1 -CsvFilePath "users.csv" -Verbose
```

## Sample CSV Files

### Basic Format
```csv
UserPrincipalName,DisplayName
john.doe@company.com,John Doe
jane.smith@company.com,Jane Smith
mike.jones@company.com,Mike Jones
```

### Alternative Format
```csv
Email,Name
john.doe@company.com,John Doe
jane.smith@company.com,Jane Smith
mike.jones@company.com,Mike Jones
```

## License and Support

This script is provided as-is for educational and compliance purposes. For production use, thoroughly test in a non-production environment first.

## Version History

- **v1.0**: Initial release with Teams chat search functionality
- **v1.1**: Added date filtering and fallback strategies
- **v1.2**: Enhanced error handling and user validation

---

**Note**: This script requires appropriate Microsoft 365 licenses and permissions. Always test in a non-production environment before using in production.
