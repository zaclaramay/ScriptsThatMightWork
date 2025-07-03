# Microsoft Purview Chat History Search Script

A PowerShell script that creates a single Microsoft Purview content search to find chat history across multiple users' mailboxes simultaneously.

## Overview

This script automates the process of searching for chat messages, emails, and other communication data across multiple Microsoft 365 users using Microsoft Purview eDiscovery. Instead of creating individual searches for each user, it creates one comprehensive search that includes all specified users' mailboxes from a CSV.

## What the Script Does

1. **Connects to Microsoft Services**: Authenticates with Microsoft Graph and Security & Compliance Center
2. **Processes User List**: Reads a CSV file containing user email addresses and validates them
3. **Creates Single Search**: Generates one content search that targets all users' mailboxes
4. **Monitors Progress**: Tracks the search execution and reports status updates
5. **Provides Results**: Displays search statistics and next steps for data export

## Supported Content Types

- **Microsoft Teams** chat messages and channel conversations
- **Yammer** messages and posts
- **Skype for Business** messages
- **Exchange Email** (optional)

## Prerequisites

### Required Software
- Windows PowerShell 5.1 or PowerShell 7+
- Microsoft 365 tenant with Purview/eDiscovery licensing

### Required PowerShell Modules
Install these modules as Administrator:
```powershell
Install-Module -Name ExchangeOnlineManagement -Force
Install-Module -Name Microsoft.Graph.Authentication -Force
Install-Module -Name Microsoft.Graph.Compliance -Force
```

### Required Permissions
You need one of these permission setups:

#### Option 1: Interactive Authentication (Recommended)
- **Global Administrator** or **Compliance Administrator** role in Microsoft 365
- The script will prompt for credentials during execution

#### Option 2: Service Principal Authentication
- Azure AD application with these API permissions:
  - **Microsoft Graph**: `SecurityEvents.Read.All`, `AuditLog.Read.All`, `Directory.Read.All`
  - **Exchange Online**: `Compliance.ReadWrite`

## CSV File Format

Create a CSV file with the following structure:

```csv
UserPrincipalName,DisplayName
john.doe@company.com,John Doe
jane.smith@company.com,Jane Smith
bob.jones@company.com,Bob Jones
mary.wilson@company.com,Mary Wilson
```

### Required Columns
- **UserPrincipalName**: User's email address (required)
- **DisplayName**: User's display name (optional)

### Alternative Column Names
The script can also recognize these column variations:
- `Email` (instead of UserPrincipalName)
- `UPN` (instead of UserPrincipalName)
- `Name` (instead of DisplayName)

## Usage

### Basic Usage
```powershell
.\PurviewChatSearch.ps1 -CsvFilePath "C:\path\to\users.csv"
```

### Specify Content Types
```powershell
.\PurviewChatSearch.ps1 -CsvFilePath "C:\users.csv" -IncludeTeamsChats -IncludeEmail
```

### Use Service Principal Authentication
```powershell
.\PurviewChatSearch.ps1 -CsvFilePath "C:\users.csv" -TenantId "your-tenant-id" -ClientId "your-app-id" -ClientSecret "your-client-secret"
```

### Custom Search Name
```powershell
.\PurviewChatSearch.ps1 -CsvFilePath "C:\users.csv" -SearchName "Q1_2024_Investigation"
```

### All Parameters
```powershell
.\PurviewChatSearch.ps1 `
    -CsvFilePath "C:\investigation\users.csv" `
    -IncludeTeamsChats `
    -IncludeYammerMessages `
    -IncludeSkypeMessages `
    -IncludeEmail `
    -SearchName "ComprehensiveSearch_2024" `
    -TenantId "12345678-1234-1234-1234-123456789012" `
    -ClientId "87654321-4321-4321-4321-210987654321" `
    -ClientSecret "your-client-secret"
```

## Parameters

| Parameter | Required | Type | Description |
|-----------|----------|------|-------------|
| `CsvFilePath` | Yes | String | Path to CSV file containing user list |
| `TenantId` | No | String | Azure AD Tenant ID (for service principal auth) |
| `ClientId` | No | String | Application ID (for service principal auth) |
| `ClientSecret` | No | String | Client secret (for service principal auth) |
| `IncludeTeamsChats` | No | Switch | Include Microsoft Teams chat messages |
| `IncludeYammerMessages` | No | Switch | Include Yammer messages |
| `IncludeSkypeMessages` | No | Switch | Include Skype for Business messages |
| `IncludeEmail` | No | Switch | Include email messages |
| `SearchName` | No | String | Custom name for the search (auto-generated if not provided) |

## Step-by-Step Execution

1. **Prepare Your Environment**
   ```powershell
   # Run PowerShell as Administrator
   # Install required modules
   Install-Module -Name ExchangeOnlineManagement -Force
   Install-Module -Name Microsoft.Graph.Authentication -Force
   Install-Module -Name Microsoft.Graph.Compliance -Force
   ```

2. **Create Your CSV File**
   - Use the format shown above
   - Ensure email addresses are valid
   - Save with UTF-8 encoding

3. **Run the Script**
   ```powershell
   # Navigate to script directory
   cd "C:\path\to\script"
   
   # Execute the script
   .\PurviewChatSearch.ps1 -CsvFilePath "C:\path\to\users.csv"
   ```

4. **Follow Authentication Prompts**
   - Enter your Microsoft 365 credentials when prompted
   - Complete any multi-factor authentication requirements

5. **Monitor Progress**
   - The script will show real-time progress updates
   - Large searches may take 30 minutes to 2 hours to complete

## Script Output

The script provides detailed output including:

- **Connection Status**: Confirmation of successful authentication
- **User Validation**: List of valid users loaded from CSV
- **Search Creation**: Confirmation of search creation with details
- **Progress Updates**: Real-time status updates during search execution
- **Final Results**: Summary statistics and next steps

### Example Output
```
=== Purview Bulk Chat History Search Script ===
Started at: 03/15/2024 10:30:00 AM

Required modules loaded successfully
Connected to Microsoft Graph successfully
Connected to Security & Compliance Center successfully
Loaded 25 valid users from CSV
Valid users to include in search: 25

--- Creating bulk search: BulkChatHistorySearch_20240315_103000 ---
Created content search: BulkChatHistorySearch_20240315_103000
Started search: BulkChatHistorySearch_20240315_103000

Search status: InProgress - Elapsed: 120s - Items found: 1,247 - Size: 45.2 MB
Search status: InProgress - Elapsed: 240s - Items found: 2,891 - Size: 112.7 MB
Search status: Completed - Elapsed: 420s - Items found: 4,156 - Size: 187.3 MB

Search completed successfully!

=== Search Results Summary ===
Search Name: BulkChatHistorySearch_20240315_103000
Status: Completed
Total Items Found: 4,156
Total Size: 187.3 MB
Users Searched: 25
Content Types: Teams Chats, Yammer Messages, Skype Messages
```

## After Script Completion

1. **Access Microsoft Purview Compliance Center**
   - Go to [https://compliance.microsoft.com](https://compliance.microsoft.com)
   - Navigate to **Content Search**

2. **Find Your Search**
   - Look for the search name displayed in the script output
   - Review search statistics and results

3. **Create Export (if needed)**
   - Click on your search
   - Select **Export results**
   - Choose export format and options
   - Download using the Microsoft Office 365 eDiscovery Export Tool

## Troubleshooting

### Common Issues and Solutions

#### Module Import Errors
```
Error: Failed to import required modules
```
**Solution**: Install modules as Administrator:
```powershell
Install-Module -Name ExchangeOnlineManagement -Force
Install-Module -Name Microsoft.Graph.Authentication -Force
Install-Module -Name Microsoft.Graph.Compliance -Force
```

#### Authentication Failures
```
Error: Failed to connect to Microsoft Graph
```
**Solutions**:
- Verify you have appropriate permissions
- Check if MFA is required and complete authentication
- For service principal auth, verify ClientId, ClientSecret, and TenantId

#### CSV Format Issues
```
Error: Could not find UserPrincipalName column
```
**Solutions**:
- Ensure CSV has `UserPrincipalName` column
- Use `Email` or `UPN` as alternative column names
- Check CSV encoding (should be UTF-8)

#### Search Creation Failures
```
Error: Failed to create content search
```
**Solutions**:
- Verify users exist in your tenant
- Check eDiscovery permissions
- Ensure proper licensing (E3/E5 with Compliance features)

#### Large Search Timeouts
```
Warning: Search monitoring timed out
```
**Solutions**:
- Check search status manually in Compliance Center
- Large searches (100+ users) may take several hours
- Consider breaking into smaller batches

## Best Practices

1. **Start Small**: Test with a few users before running large searches
2. **Use Specific Content Types**: Only include needed content types to improve performance
3. **Monitor Licensing**: Ensure sufficient eDiscovery licenses for your user count
4. **Regular Cleanup**: Remove old searches to avoid clutter
5. **Document Searches**: Use descriptive search names and keep records

## Security Considerations

- **Least Privilege**: Only grant necessary permissions
- **Audit Trail**: All searches are logged in the Microsoft 365 audit log
- **Data Handling**: Follow your organization's data retention policies
- **Access Control**: Limit script access to authorized personnel only

## Support and Limitations

### Script Limitations
- Maximum 1,000 users per search (Microsoft limitation)
- Search timeout after 2 hours of monitoring
- No automatic retry for failed searches
- Export functionality must be handled manually

### Microsoft 365 Limitations
- eDiscovery licensing required
- Some content types may have retention limitations
- Guest user data may not be searchable
- Deleted content may not be recoverable

### Getting Help
- Review Microsoft 365 eDiscovery documentation
- Check the Microsoft 365 Service Health dashboard

## License

This script is provided as-is for educational and administrative purposes. Ensure compliance with your organization's policies and Microsoft's terms of service when using this script.

## Version History

- **v1.0**: Initial release with bulk search functionality
- **v1.1**: Added enhanced error handling and user validation
- **v1.2**: Improved CSV parsing and alternative column support
