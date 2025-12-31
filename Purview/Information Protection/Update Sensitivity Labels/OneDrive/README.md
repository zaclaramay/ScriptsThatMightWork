# OneDrive Sensitivity Label Bulk Migration Tool

A PowerShell script for bulk updating Microsoft Information Protection (MIP) sensitivity labels across multiple OneDrive accounts.

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Part 1: Getting Sensitivity Label IDs](#part-1-getting-sensitivity-label-ids)
- [Part 2: Microsoft Entra ID App Registration](#part-2-microsoft-entra-id-app-registration)
- [Part 3: Azure Billing Account Setup for Metered APIs](#part-3-azure-billing-account-setup-for-metered-apis)
- [Part 4: Preparing the Users CSV File](#part-4-preparing-the-users-csv-file)
- [Part 5: Running the Script](#part-5-running-the-script)
- [Parameter Reference](#parameter-reference)
- [Output and Logging](#output-and-logging)
- [Troubleshooting](#troubleshooting)
- [Security Best Practices](#security-best-practices)

---

## Overview

This tool scans OneDrive accounts for files with a specific sensitivity label and updates them to a new label. It supports:

- Processing multiple OneDrive users from a CSV file
- Recursive scanning of all folders within each OneDrive
- Detailed logging of all changes
- Per-user summary reporting
- Error handling with continue-on-error option
- Automatic token refresh for long-running operations

---

## Prerequisites

Before using this script, ensure you have:

- **PowerShell 5.1 or later** (PowerShell 7+ recommended)
- **Global Administrator** or **Security Administrator** role in Microsoft 365
- **User Administrator** role for accessing user OneDrives
- An **Azure subscription** for metered API billing
- The following PowerShell modules (installed automatically by the script):
  - PnP.PowerShell (version 2.0.0 or later)

### Install Required Modules Manually (Optional)

```powershell
# Install PnP.PowerShell module
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force

# Install Exchange Online module (for label management)
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force

# Install Microsoft Graph module
Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
```

---

## Part 1: Getting Sensitivity Label IDs

Sensitivity labels are identified by GUIDs. Use these PowerShell commands to retrieve your label IDs.

### Method 1: Using Security & Compliance PowerShell

```powershell
# Connect to Security & Compliance Center
Connect-IPPSSession -UserPrincipalName admin@yourtenant.onmicrosoft.com

# Get all sensitivity labels with their GUIDs
Get-Label | Select-Object DisplayName, Name, Guid, Priority, ParentId | Format-Table -AutoSize

# Get detailed information about a specific label
Get-Label -Identity "Confidential" | Format-List *

# Export all labels to CSV for reference
Get-Label | Select-Object DisplayName, Name, Guid, Priority, ParentId, ContentType | 
    Export-Csv -Path ".\SensitivityLabels.csv" -NoTypeInformation
```

### Method 2: Using Microsoft Graph PowerShell

```powershell
# Install Microsoft Graph module if needed
Install-Module Microsoft.Graph -Scope CurrentUser

# Connect with required scopes
Connect-MgGraph -Scopes "InformationProtectionPolicy.Read"

# Get all sensitivity labels
$labels = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/security/informationProtection/sensitivityLabels"

# Display labels with IDs
$labels.value | Select-Object id, name, description | Format-Table -AutoSize

# Disconnect when done
Disconnect-MgGraph
```

### Method 3: Using Exchange Online PowerShell

```powershell
# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName admin@yourtenant.onmicrosoft.com

# Get sensitivity labels
Get-Label | Format-Table DisplayName, Guid -AutoSize
```

### Example Output

```
DisplayName              Guid
-----------              ----
Public                   xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
Internal                 yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy
Confidential             zzzzzzzz-zzzz-zzzz-zzzz-zzzzzzzzzzzz
Highly Confidential      aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa
```

---

## Part 2: Microsoft Entra ID App Registration

Create an app registration to authenticate the script with Microsoft Graph.

### Step 2.1: Create the App Registration

1. Navigate to **Azure Portal** → [portal.azure.com](https://portal.azure.com)
2. Go to **Microsoft Entra ID** → **App registrations**
3. Click **+ New registration**
4. Configure the application:
   - **Name**: `OneDrive Label Migration Tool`
   - **Supported account types**: "Accounts in this organizational directory only"
   - **Redirect URI**: Leave blank
5. Click **Register**

### Step 2.2: Record Application IDs

After registration, copy these values from the **Overview** page:

| Field | Description | Script Parameter |
|-------|-------------|------------------|
| Application (client) ID | Unique app identifier | `-ClientId` |
| Directory (tenant) ID | Your Azure AD tenant | `-TenantId` |

### Step 2.3: Create a Client Secret

1. In your app registration, go to **Certificates & secrets**
2. Click **+ New client secret**
3. Configure:
   - **Description**: `OneDrive Label Migration Script`
   - **Expires**: Choose 12 or 24 months
4. Click **Add**
5. **IMMEDIATELY copy the Value** (shown only once) → This is your `-ClientSecret`

### Step 2.4: Configure API Permissions

1. Go to **API permissions** → **+ Add a permission**

2. **Add Microsoft Graph permissions**:
   - Click **Microsoft Graph** → **Application permissions**
   - Add these permissions:

   | Permission | Purpose |
   |------------|---------|
   | `Files.ReadWrite.All` | Read and write all users' OneDrive files |
   | `User.Read.All` | Read user profiles to access OneDrive |
   | `InformationProtectionPolicy.Read.All` | Read sensitivity label policies |

### Step 2.5: Grant Admin Consent

1. After adding all permissions, click **Grant admin consent for [Your Organization]**
2. Click **Yes** to confirm
3. Verify all permissions show a green checkmark ✓ under "Status"

### Final Permissions Summary

```
Microsoft Graph (Application)
├── Files.ReadWrite.All                    ✓ Granted
├── User.Read.All                          ✓ Granted
└── InformationProtectionPolicy.Read.All   ✓ Granted
```

---

## Part 3: Azure Billing Account Setup for Metered APIs

The Microsoft Graph sensitivity label APIs (`assignSensitivityLabel`, `extractSensitivityLabels`) are **metered APIs** that require billing configuration.

### Step 3.1: Understand Metered API Pricing

- Metered APIs charge per API call beyond a free threshold
- Current pricing: ~$0.00025 per call (verify at [Microsoft Graph pricing](https://learn.microsoft.com/en-us/graph/metered-api-overview))
- Free tier: First 1,000 calls/month per tenant

### Step 3.2: Enable Metered API Billing

1. **Navigate to Azure Portal** → [portal.azure.com](https://portal.azure.com)

2. **Create or Select a Resource Group**:
   ```
   Azure Portal → Resource groups → + Create
   - Subscription: [Your subscription]
   - Resource group name: rg-graph-metered-apis
   - Region: [Your region]
   ```

3. **Enable Graph API Billing**:
   
   a. Go to **Microsoft Entra ID** → **App registrations** → Select your app
   
   b. Go to **API permissions**
   
   c. For each metered API permission, you need to link it to a billing subscription

4. **Link to Azure Subscription via PowerShell**:

   ```powershell
   # Install Azure PowerShell module if needed
   Install-Module -Name Az -Scope CurrentUser -Force
   
   # Connect to Azure
   Connect-AzAccount
   
   # Set context to your subscription
   Set-AzContext -SubscriptionId "your-subscription-id"
   
   # Register the Microsoft.GraphServices resource provider
   Register-AzResourceProvider -ProviderNamespace Microsoft.GraphServices
   ```

5. **Create Graph Services Account** (via Azure Portal):
   
   a. Search for "Graph Services" in Azure Portal
   
   b. Click **+ Create**
   
   c. Configure:
      - **Subscription**: Your Azure subscription
      - **Resource group**: rg-graph-metered-apis
      - **Name**: graph-metered-billing
      - **Application ID**: Paste your App Registration Client ID
   
   d. Click **Review + create** → **Create**

### Step 3.3: Verify Billing Setup

```powershell
# Test the API by making a simple call
$token = Get-GraphAccessToken -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret

$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type" = "application/json"
}

# This should work without "paymentRequired" error
Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users" -Headers $headers
```

### Alternative: Azure CLI Method

```bash
# Login to Azure
az login

# Register the resource provider
az provider register --namespace Microsoft.GraphServices

# Create the Graph Services account
az resource create \
    --resource-group rg-graph-metered-apis \
    --resource-type Microsoft.GraphServices/accounts \
    --name graph-metered-billing \
    --properties '{"appId": "your-client-id"}'
```

---

## Part 4: Preparing the Users CSV File

Create a CSV file listing all OneDrive users to process.

### CSV Format

```csv
UserPrincipalName
john.doe@contoso.com
jane.smith@contoso.com
bob.wilson@contoso.com
alice.johnson@contoso.com
```

### Supported Column Names

| Column | Description |
|--------|-------------|
| `UserPrincipalName` | User's UPN (email address) - **Recommended** |
| `Email` | Alternative to UserPrincipalName |
| `UserId` | Azure AD User Object ID |

### Generate User List from Microsoft Graph

```powershell
# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All"

# Export all licensed users to CSV
Get-MgUser -All -Filter "assignedLicenses/`$count ne 0" | 
    Select-Object @{N='UserPrincipalName';E={$_.UserPrincipalName}} |
    Export-Csv -Path ".\users.csv" -NoTypeInformation

# Or filter by department
Get-MgUser -All -Filter "department eq 'Engineering'" | 
    Select-Object @{N='UserPrincipalName';E={$_.UserPrincipalName}} |
    Export-Csv -Path ".\users.csv" -NoTypeInformation

# Disconnect when done
Disconnect-MgGraph
```

### Generate User List from Azure AD PowerShell

```powershell
# Connect to Azure AD
Connect-AzureAD

# Export all users with OneDrive
Get-AzureADUser -All $true | 
    Where-Object { $_.AssignedLicenses.Count -gt 0 } |
    Select-Object @{N='UserPrincipalName';E={$_.UserPrincipalName}} |
    Export-Csv -Path ".\users.csv" -NoTypeInformation
```

### Generate User List from Exchange Online

```powershell
# Connect to Exchange Online
Connect-ExchangeOnline

# Get all mailbox users
Get-Mailbox -ResultSize Unlimited | 
    Select-Object @{N='UserPrincipalName';E={$_.UserPrincipalName}} |
    Export-Csv -Path ".\users.csv" -NoTypeInformation
```

---

## Part 5: Running the Script

### Basic Usage

```powershell
.\Update-OneDriveSensitivityLabels-Bulk.ps1 `
    -UsersCSVPath ".\users.csv" `
    -SourceLabelId "your-source-label-guid" `
    -TargetLabelId "your-target-label-guid" `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -ClientSecret "your-client-secret"
```

### Test Run (Limited Users)

```powershell
# Process only first 3 users for testing
.\Update-OneDriveSensitivityLabels-Bulk.ps1 `
    -UsersCSVPath ".\users.csv" `
    -SourceLabelId "your-source-label-guid" `
    -TargetLabelId "your-target-label-guid" `
    -ClientId "your-client-id" `
    -TenantId "your-tenant-id" `
    -ClientSecret "your-secret" `
    -MaxConcurrentUsers 3
```

### Continue on Errors

```powershell
# Don't stop if one user's OneDrive fails
.\Update-OneDriveSensitivityLabels-Bulk.ps1 `
    -UsersCSVPath ".\users.csv" `
    -SourceLabelId "your-source-label-guid" `
    -TargetLabelId "your-target-label-guid" `
    -ClientId "your-client-id" `
    -TenantId "your-tenant-id" `
    -ClientSecret "your-secret" `
    -ContinueOnError
```

### Custom Log Paths

```powershell
.\Update-OneDriveSensitivityLabels-Bulk.ps1 `
    -UsersCSVPath ".\users.csv" `
    -SourceLabelId "your-source-label-guid" `
    -TargetLabelId "your-target-label-guid" `
    -ClientId "your-client-id" `
    -TenantId "your-tenant-id" `
    -ClientSecret "your-secret" `
    -LogPath "C:\Logs\OneDriveDetailedChanges.csv" `
    -SummaryLogPath "C:\Logs\OneDriveUserSummary.csv"
```

### Full Example with All Parameters

```powershell
.\Update-OneDriveSensitivityLabels-Bulk.ps1 `
    -UsersCSVPath "C:\Migration\users.csv" `
    -SourceLabelId "your-source-label-guid" `
    -TargetLabelId "your-target-label-guid" `
    -LogPath "C:\Migration\Logs\DetailLog_$(Get-Date -Format 'yyyyMMdd').csv" `
    -SummaryLogPath "C:\Migration\Logs\Summary_$(Get-Date -Format 'yyyyMMdd').csv" `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -ClientSecret "your-secret-here" `
    -MaxConcurrentUsers 0 `
    -ContinueOnError
```

---

## Parameter Reference

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-UsersCSVPath` | Yes | - | Path to CSV file containing OneDrive users |
| `-SourceLabelId` | Yes | - | GUID of sensitivity label to search for |
| `-TargetLabelId` | Yes | - | GUID of sensitivity label to apply |
| `-LogPath` | No | `.\OneDriveLabelChanges_[timestamp].csv` | Path for detailed change log |
| `-SummaryLogPath` | No | `.\OneDriveLabelSummary_[timestamp].csv` | Path for per-user summary log |
| `-ClientId` | Yes | - | Azure AD App Registration Client ID |
| `-TenantId` | Yes | - | Azure AD Tenant ID |
| `-ClientSecret` | Yes | - | Azure AD App Client Secret |
| `-MaxConcurrentUsers` | No | `0` (all) | Limit number of users to process |
| `-ContinueOnError` | No | `$false` | Continue if a user fails |

---

## Output and Logging

### Console Output

The script displays real-time progress:

```
==========================================
OneDrive Bulk Sensitivity Label Migration
==========================================

Reading users from: .\users.csv
Processing 5 users
Source Label: your-source-label-guid
Target Label: your-target-label-guid

[1/5] Processing: john.doe@contoso.com
  OneDrive found: https://contoso-my.sharepoint.com/personal/john_doe_contoso_com
  Found 150 files
    Updated: Budget_2024.xlsx
    Updated: Project_Plan.docx
  Completed: 2 updated, 0 failed

[2/5] Processing: jane.smith@contoso.com
...
```

### Detail Log (CSV)

Records every file change:

| Column | Description |
|--------|-------------|
| Timestamp | When the change occurred |
| UserPrincipalName | User's email/UPN |
| UserId | Azure AD Object ID |
| FileName | Name of the file |
| FileUrl | Full URL to the file |
| FilePath | Path within OneDrive |
| OriginalLabelId | Previous label GUID |
| OriginalLabelName | Previous label display name |
| NewLabelId | New label GUID |
| NewLabelName | New label display name |
| Status | Updated, Failed, or Error |
| ErrorMessage | Error details if failed |
| FileSize | File size in bytes |
| LastModified | Last modification date |

### Summary Log (CSV)

Records per-user statistics:

| Column | Description |
|--------|-------------|
| Timestamp | Processing timestamp |
| UserPrincipalName | User's email/UPN |
| UserId | Azure AD Object ID |
| TotalFiles | Total files in OneDrive |
| FilesScanned | Files scanned for labels |
| FilesMatched | Files with source label |
| FilesUpdated | Successfully updated |
| FilesFailed | Failed to update |
| Status | Completed or Error |
| ErrorMessage | Error details if failed |
| ProcessingTime | Time to process user |

---

## Troubleshooting

### Common Errors and Solutions

| Error | Cause | Solution |
|-------|-------|----------|
| `paymentRequired` | Metered API billing not configured | Complete [Part 3](#part-3-azure-billing-account-setup-for-metered-apis) |
| `Insufficient privileges` | Missing API permissions | Verify permissions and admin consent in [Part 2](#part-2-microsoft-entra-id-app-registration) |
| `User not found` | Invalid UPN or user doesn't exist | Verify user exists in Azure AD |
| `Drive not found` | User doesn't have OneDrive provisioned | Ensure user has OneDrive license and has accessed OneDrive at least once |
| `Invalid client secret` | Secret expired or wrong | Create a new secret in App Registration |
| `Token refresh failed` | Network or auth issue | Re-run script; check credentials |

### Verify App Permissions

```powershell
# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.Read.All"

# Get your app and its permissions
$app = Get-MgApplication -Filter "appId eq 'your-client-id'"
$app.RequiredResourceAccess | ForEach-Object {
    $resourceId = $_.ResourceAppId
    $_.ResourceAccess | Select-Object @{N='Resource';E={$resourceId}}, Id, Type
}
```

### Test OneDrive Access

```powershell
# Quick test to verify OneDrive access
$token = # ... get token using your credentials
$headers = @{ "Authorization" = "Bearer $token" }

$userUpn = "john.doe@contoso.com"
$graphUrl = "https://graph.microsoft.com/v1.0/users/$userUpn/drive"

Invoke-RestMethod -Uri $graphUrl -Headers $headers
```

### Check if User Has OneDrive Provisioned

```powershell
# Connect to SharePoint Online
Connect-SPOService -Url https://contoso-admin.sharepoint.com

# Check specific user's OneDrive
$userUpn = "john.doe@contoso.com"
$oneDriveUrl = "https://contoso-my.sharepoint.com/personal/" + ($userUpn -replace '[@.]', '_')

try {
    Get-SPOSite -Identity $oneDriveUrl
    Write-Host "OneDrive exists for $userUpn"
} catch {
    Write-Host "OneDrive NOT provisioned for $userUpn"
}
```

### Enable Verbose Logging

```powershell
# Run with verbose output for debugging
.\Update-OneDriveSensitivityLabels-Bulk.ps1 `
    -UsersCSVPath ".\users.csv" `
    -SourceLabelId "your-source-label-guid" `
    -TargetLabelId "your-target-label-guid" `
    -ClientId "your-client-id" `
    -TenantId "your-tenant-id" `
    -ClientSecret "your-secret" `
    -Verbose
```

---

## Security Best Practices

### Protect Client Secrets

```powershell
# Option 1: Use Windows Credential Manager
$credential = Get-Credential -UserName $ClientId -Message "Enter Client Secret"
$ClientSecret = $credential.GetNetworkCredential().Password

# Option 2: Use Azure Key Vault
$secret = Get-AzKeyVaultSecret -VaultName "YourVault" -Name "GraphClientSecret"
$ClientSecret = $secret.SecretValueText

# Option 3: Use environment variables
$env:GRAPH_CLIENT_SECRET = "your-secret"
$ClientSecret = $env:GRAPH_CLIENT_SECRET
```

### Use Certificate Authentication (Recommended for Production)

```powershell
# Generate a self-signed certificate
$cert = New-SelfSignedCertificate `
    -Subject "CN=OneDriveLabelMigration" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -NotAfter (Get-Date).AddYears(2)

# Export certificate for upload to Azure AD
Export-Certificate -Cert $cert -FilePath ".\OneDriveLabelMigration.cer"

# Upload .cer file to App Registration → Certificates & secrets → Certificates
```

### Limit Scope to Specific Users

For enhanced security, consider processing users in batches based on department or role:

```powershell
# Get users by department
Get-MgUser -All -Filter "department eq 'Finance'" | 
    Select-Object @{N='UserPrincipalName';E={$_.UserPrincipalName}} |
    Export-Csv -Path ".\finance-users.csv" -NoTypeInformation
```

### Audit and Monitor

- Enable **Azure AD Sign-in logs** for the app
- Set up **Alerts** for unusual activity
- Review **Summary logs** after each run
- Rotate client secrets regularly (every 6-12 months)

---

## Comparison: OneDrive vs SharePoint Script

| Feature | OneDrive Script | SharePoint Script |
|---------|-----------------|-------------------|
| Input CSV | User list (UPN/Email) | Site URLs |
| Scope | Per-user OneDrive | Per-site document library |
| API Endpoint | `/users/{id}/drive` | `/sites/{id}/drives` |
| Best for | User-centric migrations | Site-centric migrations |

---

## Support

For issues or questions:

1. Check the [Troubleshooting](#troubleshooting) section
2. Review the generated log files
3. Verify all prerequisites are met
4. Test with a single user using `-MaxConcurrentUsers 1`

---

## License

This script is provided as-is without warranty. Use at your own risk. Always test in a non-production environment first.
