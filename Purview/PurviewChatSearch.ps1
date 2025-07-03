# Purview Chat Data Search Script - Single Search for All Users
# This script creates a single content search for all users listed in a CSV file using Microsoft Purview

param(
    [Parameter(Mandatory = $true)]
    [string]$CsvFilePath,
    
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeTeamsChats,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeYammerMessages,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSkypeMessages,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeEmail,
    
    [Parameter(Mandatory = $false)]
    [string]$SearchName
)

# Import required modules
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Compliance -ErrorAction Stop
    Write-Host "Required modules loaded successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to import required modules. Please install: ExchangeOnlineManagement, Microsoft.Graph.Authentication, Microsoft.Graph.Compliance"
    exit 1
}

# Function to authenticate to Microsoft Graph
function Connect-ToGraph {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )
    
    try {
        if ($ClientId -and $ClientSecret -and $TenantId) {
            $SecureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
            $Credential = New-Object System.Management.Automation.PSCredential($ClientId, $SecureSecret)
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $Credential -Scopes "https://graph.microsoft.com/.default"
        }
        else {
            Connect-MgGraph -Scopes "SecurityEvents.Read.All", "AuditLog.Read.All", "Directory.Read.All"
        }
        Write-Host "Connected to Microsoft Graph successfully" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
        exit 1
    }
}

# Function to connect to Security & Compliance Center
function Connect-ToComplianceCenter {
    try {
        Connect-IPPSSession
        Write-Host "Connected to Security & Compliance Center successfully" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Security & Compliance Center: $($_.Exception.Message)"
        exit 1
    }
}

# Function to create content search for all users
function New-BulkContentSearch {
    param(
        [string]$SearchName,
        [array]$UserPrincipalNames
    )
    
    # Build the content type filters
    $KqlQuery = @()
    
    if ($IncludeTeamsChats) {
        $KqlQuery += "kind:microsoftteams"
    }
    
    if ($IncludeYammerMessages) {
        $KqlQuery += "kind:yammer"
    }
    
    if ($IncludeSkypeMessages) {
        $KqlQuery += "kind:skypeforbusiness"
    }
    
    if ($IncludeEmail) {
        $KqlQuery += "kind:email"
    }
    
    # Default to all chat types if none specified
    if ($KqlQuery.Count -eq 0) {
        $KqlQuery += "kind:microsoftteams"
        $KqlQuery += "kind:yammer"
        $KqlQuery += "kind:skypeforbusiness"
    }
    
    # Create query for all chat history (no date restrictions)
    $Query = "(" + ($KqlQuery -join " OR ") + ")"
    
    Write-Host "KQL Query: $Query" -ForegroundColor Cyan
    Write-Host "Searching all chat history for $($UserPrincipalNames.Count) users" -ForegroundColor Cyan
    
    try {
        $SearchParams = @{
            Name = $SearchName
            ExchangeLocation = $UserPrincipalNames  # Target all users' mailboxes
            ContentMatchQuery = $Query
            Description = "Complete chat history search for $($UserPrincipalNames.Count) users - Created $(Get-Date)"
        }
        
        New-ComplianceSearch @SearchParams
        Write-Host "Created content search: $SearchName" -ForegroundColor Green
        Write-Host "Query used: $Query" -ForegroundColor Gray
        Write-Host "Target mailboxes: $($UserPrincipalNames.Count) users" -ForegroundColor Gray
        return $true
    }
    catch {
        Write-Error "Failed to create content search: $($_.Exception.Message)"
        Write-Host "Query that failed: $Query" -ForegroundColor Red
        
        # Try fallback with simpler search
        Write-Host "Attempting fallback search..." -ForegroundColor Yellow
        return New-BulkContentSearchFallback -SearchName $SearchName -UserPrincipalNames $UserPrincipalNames
    }
}

# Fallback function for content search
function New-BulkContentSearchFallback {
    param(
        [string]$SearchName,
        [array]$UserPrincipalNames
    )
    
    try {
        # Very simple fallback query - all Teams chats
        $Query = "kind:microsoftteams"
        
        Write-Host "Fallback KQL Query: $Query" -ForegroundColor Yellow
        
        $SearchParams = @{
            Name = "$SearchName" + "_Fallback"
            ExchangeLocation = $UserPrincipalNames
            ContentMatchQuery = $Query
            Description = "Fallback complete chat search for $($UserPrincipalNames.Count) users - Created $(Get-Date)"
        }
        
        New-ComplianceSearch @SearchParams
        Write-Host "Created fallback content search: $($SearchName)_Fallback" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Fallback search also failed: $($_.Exception.Message)"
        return $false
    }
}

# Function to start and monitor search
function Start-AndMonitorSearch {
    param(
        [string]$SearchName
    )
    
    try {
        # Verify search exists before starting
        $Search = Get-ComplianceSearch -Identity $SearchName -ErrorAction SilentlyContinue
        if (-not $Search) {
            Write-Error "Search not found: $SearchName"
            return $false
        }
        
        Start-ComplianceSearch -Identity $SearchName
        Write-Host "Started search: $SearchName" -ForegroundColor Yellow
        
        $MaxWaitTime = 7200  # 2 hours max wait time for large bulk searches
        $WaitTime = 0
        
        do {
            Start-Sleep -Seconds 120  # Check every 2 minutes for large searches
            $WaitTime += 120
            
            $SearchStatus = Get-ComplianceSearch -Identity $SearchName -ErrorAction SilentlyContinue
            if (-not $SearchStatus) {
                Write-Error "Could not retrieve search status for: $SearchName"
                return $false
            }
            
            Write-Host "Search status: $($SearchStatus.Status) - Elapsed: $($WaitTime)s - Items found: $($SearchStatus.Items) - Size: $($SearchStatus.Size)" -ForegroundColor Cyan
            
            if ($WaitTime -ge $MaxWaitTime) {
                Write-Warning "Search monitoring timed out after $MaxWaitTime seconds"
                Write-Warning "You may need to manually check the search status in the Compliance Center"
                return $false
            }
            
        } while ($SearchStatus.Status -eq "InProgress" -or $SearchStatus.Status -eq "Starting")
        
        if ($SearchStatus.Status -eq "Completed") {
            Write-Host "Search completed successfully: $SearchName" -ForegroundColor Green
            Write-Host "Total items found: $($SearchStatus.Items)" -ForegroundColor Green
            Write-Host "Total size: $($SearchStatus.Size)" -ForegroundColor Green
            return $true
        }
        elseif ($SearchStatus.Status -eq "CompletedWithErrors") {
            Write-Warning "Search completed with errors: $SearchName"
            Write-Host "Items found: $($SearchStatus.Items)" -ForegroundColor Yellow
            Write-Host "Errors: $($SearchStatus.Errors)" -ForegroundColor Red
            return $true  # Still return true as we have some results
        }
        else {
            Write-Error "Search failed with status: $($SearchStatus.Status)"
            if ($SearchStatus.Errors) {
                Write-Error "Errors: $($SearchStatus.Errors)"
            }
            return $false
        }
    }
    catch {
        Write-Error "Failed to start or monitor search: $($_.Exception.Message)"
        return $false
    }
}

# Function to process user list from CSV
function Get-UsersFromCsv {
    param([string]$CsvPath)
    
    try {
        if (-not (Test-Path $CsvPath)) {
            throw "CSV file not found: $CsvPath"
        }
        
        $Users = Import-Csv $CsvPath
        
        if ($Users.Count -eq 0) {
            throw "CSV file is empty or has no data rows"
        }
        
        # Validate CSV has required columns
        $RequiredColumns = @("UserPrincipalName", "DisplayName")
        $CsvColumns = $Users[0].PSObject.Properties.Name
        
        $MissingColumns = $RequiredColumns | Where-Object { $_ -notin $CsvColumns }
        if ($MissingColumns) {
            Write-Warning "Missing columns in CSV: $($MissingColumns -join ', ')"
            Write-Host "Expected columns: UserPrincipalName, DisplayName" -ForegroundColor Yellow
            Write-Host "Available columns: $($CsvColumns -join ', ')" -ForegroundColor Yellow
            
            # Try to use email column if UserPrincipalName is missing
            if ("Email" -in $CsvColumns -and "UserPrincipalName" -notin $CsvColumns) {
                Write-Host "Using 'Email' column as UserPrincipalName" -ForegroundColor Yellow
                $Users = $Users | Select-Object @{Name="UserPrincipalName";Expression={$_.Email}}, 
                                             @{Name="DisplayName";Expression={if($_.DisplayName){$_.DisplayName}elseif($_.Name){$_.Name}else{$_.Email}}}
            }
            elseif ("UPN" -in $CsvColumns -and "UserPrincipalName" -notin $CsvColumns) {
                Write-Host "Using 'UPN' column as UserPrincipalName" -ForegroundColor Yellow
                $Users = $Users | Select-Object @{Name="UserPrincipalName";Expression={$_.UPN}}, 
                                             @{Name="DisplayName";Expression={if($_.DisplayName){$_.DisplayName}elseif($_.Name){$_.Name}else{$_.UPN}}}
            }
            else {
                throw "Could not find UserPrincipalName, Email, or UPN column in CSV"
            }
        }
        
        # Filter out empty rows
        $Users = $Users | Where-Object { 
            $_.UserPrincipalName -and 
            $_.UserPrincipalName.Trim() -ne "" -and
            $_.UserPrincipalName -match "@"
        }
        
        if ($Users.Count -eq 0) {
            throw "No valid users found in CSV after filtering"
        }
        
        Write-Host "Loaded $($Users.Count) valid users from CSV" -ForegroundColor Green
        return $Users
    }
    catch {
        Write-Error "Failed to process CSV file: $($_.Exception.Message)"
        exit 1
    }
}

# Function to get search statistics
function Get-SearchStatistics {
    param([string]$SearchName)
    
    try {
        $Search = Get-ComplianceSearch -Identity $SearchName -ErrorAction SilentlyContinue
        if (-not $Search) {
            Write-Warning "Search not found: $SearchName"
            return $null
        }
        
        $SearchStats = @{
            Name = $Search.Name
            Status = $Search.Status
            Items = $Search.Items
            Size = $Search.Size
            ContentMatchQuery = $Search.ContentMatchQuery
            ExchangeLocationCount = $Search.ExchangeLocation.Count
            ExchangeLocations = $Search.ExchangeLocation
            Description = $Search.Description
            CreatedTime = $Search.CreatedTime
            LastModifiedTime = $Search.LastModifiedTime
            RunBy = $Search.RunBy
        }
        
        return $SearchStats
    }
    catch {
        Write-Error "Failed to get search statistics: $($_.Exception.Message)"
        return $null
    }
}

# Main execution
Write-Host "=== Purview Bulk Chat History Search Script ===" -ForegroundColor Cyan
Write-Host "Started at: $(Get-Date)" -ForegroundColor Gray
Write-Host "WARNING: This will search ALL chat history for all users in a single search" -ForegroundColor Red

# Validate parameters
if (-not (Test-Path $CsvFilePath)) {
    Write-Error "CSV file not found: $CsvFilePath"
    exit 1
}

# Show what will be included
$IncludedTypes = @()
if ($IncludeTeamsChats) { $IncludedTypes += "Teams Chats" }
if ($IncludeYammerMessages) { $IncludedTypes += "Yammer Messages" }
if ($IncludeSkypeMessages) { $IncludedTypes += "Skype Messages" }
if ($IncludeEmail) { $IncludedTypes += "Email" }

if ($IncludedTypes.Count -eq 0) {
    $IncludedTypes = @("Teams Chats", "Yammer Messages", "Skype Messages")
    Write-Host "No specific content types specified - including: $($IncludedTypes -join ', ')" -ForegroundColor Gray
}
else {
    Write-Host "Including content types: $($IncludedTypes -join ', ')" -ForegroundColor Gray
}

# Connect to services
Connect-ToGraph -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
Connect-ToComplianceCenter

# Load users from CSV
$Users = Get-UsersFromCsv -CsvPath $CsvFilePath

# Validate and prepare user list
$ValidUsers = @()
foreach ($User in $Users) {
    if (-not $User.UserPrincipalName -or $User.UserPrincipalName.Trim() -eq "") {
        Write-Warning "Skipping user with empty UserPrincipalName"
        continue
    }
    
    $UserPrincipalName = $User.UserPrincipalName.Trim()
    
    # Validate email format
    if ($UserPrincipalName -notmatch "^[^@]+@[^@]+\.[^@]+$") {
        Write-Warning "Skipping user with invalid email format: $UserPrincipalName"
        continue
    }
    
    $ValidUsers += $UserPrincipalName
}

if ($ValidUsers.Count -eq 0) {
    Write-Error "No valid users found in CSV file"
    exit 1
}

Write-Host "Valid users to include in search: $($ValidUsers.Count)" -ForegroundColor Green

# Generate search name if not provided
if (-not $SearchName) {
    $SearchName = "BulkChatHistorySearch_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
}

Write-Host "`n--- Creating bulk search: $SearchName ---" -ForegroundColor Yellow
Write-Host "This will search ALL chat history for $($ValidUsers.Count) users" -ForegroundColor Yellow

try {
    # Create content search for all users
    $SearchCreated = New-BulkContentSearch -SearchName $SearchName -UserPrincipalNames $ValidUsers
    
    if ($SearchCreated) {
        Write-Host "`nSearch created successfully. Starting search..." -ForegroundColor Green
        
        # Start and monitor search
        $SearchCompleted = Start-AndMonitorSearch -SearchName $SearchName
        
        if ($SearchCompleted) {
            Write-Host "`nSearch completed successfully!" -ForegroundColor Green
            
            # Get final search statistics
            $SearchStats = Get-SearchStatistics -SearchName $SearchName
            
            if ($SearchStats) {
                Write-Host "`n=== Search Results Summary ===" -ForegroundColor Cyan
                Write-Host "Search Name: $($SearchStats.Name)" -ForegroundColor White
                Write-Host "Status: $($SearchStats.Status)" -ForegroundColor White
                Write-Host "Total Items Found: $($SearchStats.Items)" -ForegroundColor White
                Write-Host "Total Size: $($SearchStats.Size)" -ForegroundColor White
                Write-Host "Users Searched: $($SearchStats.ExchangeLocationCount)" -ForegroundColor White
                Write-Host "Content Types: $($IncludedTypes -join ', ')" -ForegroundColor White
                Write-Host "Query Used: $($SearchStats.ContentMatchQuery)" -ForegroundColor Gray
                Write-Host "Created: $($SearchStats.CreatedTime)" -ForegroundColor Gray
                Write-Host "Run By: $($SearchStats.RunBy)" -ForegroundColor Gray
                
                # Show user list
                Write-Host "`n=== Users Included in Search ===" -ForegroundColor Cyan
                $ValidUsers | ForEach-Object { Write-Host "- $_" -ForegroundColor White }
            }
            
            Write-Host "`n=== Next Steps ===" -ForegroundColor Cyan
            Write-Host "1. Go to Security & Compliance Center (https://compliance.microsoft.com)" -ForegroundColor White
            Write-Host "2. Navigate to Content Search" -ForegroundColor White
            Write-Host "3. Find your search: $SearchName" -ForegroundColor White
            Write-Host "4. Review the results and create an export if needed" -ForegroundColor White
            Write-Host "5. Use the export functionality in the Compliance Center to download data" -ForegroundColor White
        }
        else {
            Write-Error "Search failed to complete successfully"
            exit 1
        }
    }
    else {
        Write-Error "Failed to create search"
        exit 1
    }
}
catch {
    Write-Error "Error during search execution: $($_.Exception.Message)"
    exit 1
}

Write-Host "`nScript completed at: $(Get-Date)" -ForegroundColor Gray
Write-Host "Search '$SearchName' is available in the Microsoft Purview Compliance Center" -ForegroundColor Green