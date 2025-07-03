# Purview Chat Data Export Script - All Chat History
# This script exports all chat data for users listed in a CSV file using Microsoft Purview

param(
    [Parameter(Mandatory = $true)]
    [string]$CsvFilePath,
    
    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,
    
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
    [switch]$IncludeEmail
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

# Function to create content search for all chat data
function New-ContentSearch {
    param(
        [string]$SearchName,
        [string]$UserPrincipalName
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
    Write-Host "Searching all chat history for user: $UserPrincipalName" -ForegroundColor Cyan
    
    try {
        $SearchParams = @{
            Name = $SearchName
            ExchangeLocation = @($UserPrincipalName)  # Target specific user's mailbox
            ContentMatchQuery = $Query
            Description = "Complete chat history export for user: $UserPrincipalName"
        }
        
        New-ComplianceSearch @SearchParams
        Write-Host "Created content search: $SearchName" -ForegroundColor Green
        Write-Host "Query used: $Query" -ForegroundColor Gray
        Write-Host "Target mailbox: $UserPrincipalName" -ForegroundColor Gray
        return $true
    }
    catch {
        Write-Error "Failed to create content search for $UserPrincipalName`: $($_.Exception.Message)"
        Write-Host "Query that failed: $Query" -ForegroundColor Red
        
        # Try fallback with simpler search
        Write-Host "Attempting fallback search..." -ForegroundColor Yellow
        return New-ContentSearchFallback -SearchName $SearchName -UserPrincipalName $UserPrincipalName
    }
}

# Fallback function for content search
function New-ContentSearchFallback {
    param(
        [string]$SearchName,
        [string]$UserPrincipalName
    )
    
    try {
        # Very simple fallback query - all Teams chats
        $Query = "kind:microsoftteams"
        
        Write-Host "Fallback KQL Query: $Query" -ForegroundColor Yellow
        
        $SearchParams = @{
            Name = "$SearchName" + "_Fallback"
            ExchangeLocation = @($UserPrincipalName)
            ContentMatchQuery = $Query
            Description = "Fallback complete chat export for user: $UserPrincipalName"
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
        
        $MaxWaitTime = 3600  # 60 minutes max wait time for all chat history
        $WaitTime = 0
        
        do {
            Start-Sleep -Seconds 60  # Check every minute for large searches
            $WaitTime += 60
            
            $SearchStatus = Get-ComplianceSearch -Identity $SearchName -ErrorAction SilentlyContinue
            if (-not $SearchStatus) {
                Write-Error "Could not retrieve search status for: $SearchName"
                return $false
            }
            
            Write-Host "Search status: $($SearchStatus.Status) - Elapsed: $($WaitTime)s - Items found: $($SearchStatus.Items)" -ForegroundColor Cyan
            
            if ($WaitTime -ge $MaxWaitTime) {
                Write-Warning "Search timed out after $MaxWaitTime seconds"
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

# Function to export search results
function Export-SearchResults {
    param(
        [string]$SearchName,
        [string]$ExportName,
        [string]$OutputPath
    )
    
    try {
        # Fixed: Use correct New-ComplianceSearchAction syntax
        New-ComplianceSearchAction -SearchName $SearchName -Export -ExchangeArchiveFormat "SinglePst" -SharePointArchiveFormat "SingleZip" -Format "FxStream" -IncludeCredential
        
        Write-Host "Started export: $ExportName" -ForegroundColor Yellow
        Write-Host "Note: Large chat histories may take significant time to export" -ForegroundColor Yellow
        
        $MaxWaitTime = 7200  # 2 hours max wait time for exports
        $WaitTime = 0
        
        do {
            Start-Sleep -Seconds 120  # Check every 2 minutes for large exports
            $WaitTime += 120
            
            $ExportStatus = Get-ComplianceSearchAction -Identity "$SearchName" + "_Export" -ErrorAction SilentlyContinue
            if (-not $ExportStatus) {
                Write-Warning "Could not retrieve export status for: $ExportName"
                break
            }
            
            Write-Host "Export status: $($ExportStatus.Status) - Elapsed: $($WaitTime)s" -ForegroundColor Cyan
            
            if ($WaitTime -ge $MaxWaitTime) {
                Write-Warning "Export monitoring timed out after $MaxWaitTime seconds"
                Write-Warning "Export may still be running. Check the Compliance Center for status"
                break
            }
            
        } while ($ExportStatus.Status -eq "InProgress")
        
        if ($ExportStatus.Status -eq "Completed") {
            Write-Host "Export completed successfully: $ExportName" -ForegroundColor Green
            
            # Get download URL and instructions
            $ExportDetails = Get-ComplianceSearchAction -Identity "$SearchName" + "_Export" -IncludeCredential -ErrorAction SilentlyContinue
            
            $ExportInfo = @{
                ContainerUrl = if ($ExportDetails.Results) { $ExportDetails.Results } else { "Check Compliance Center" }
                ExportName = $ExportName
                SearchName = $SearchName
                Status = "Completed"
            }
            
            return $ExportInfo
        }
        else {
            Write-Warning "Export status: $($ExportStatus.Status)"
            $ExportInfo = @{
                ExportName = $ExportName
                SearchName = $SearchName
                Status = $ExportStatus.Status
            }
            return $ExportInfo
        }
    }
    catch {
        Write-Error "Failed to export search results: $($_.Exception.Message)"
        return $null
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

# Main execution
Write-Host "=== Purview Complete Chat History Export Script ===" -ForegroundColor Cyan
Write-Host "Started at: $(Get-Date)" -ForegroundColor Gray
Write-Host "WARNING: This will export ALL chat history for each user (no date limits)" -ForegroundColor Red

# Validate parameters
if (-not (Test-Path $CsvFilePath)) {
    Write-Error "CSV file not found: $CsvFilePath"
    exit 1
}

if (-not (Test-Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    Write-Host "Created output directory: $OutputDirectory" -ForegroundColor Green
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

# Create results tracking
$Results = @()

# Process each user
foreach ($User in $Users) {
    # Validate user data
    if (-not $User.UserPrincipalName -or $User.UserPrincipalName.Trim() -eq "") {
        Write-Warning "Skipping user with empty UserPrincipalName"
        continue
    }
    
    $UserPrincipalName = $User.UserPrincipalName.Trim()
    $DisplayName = if ($User.DisplayName) { $User.DisplayName.Trim() } else { $UserPrincipalName }
    
    # Validate email format
    if ($UserPrincipalName -notmatch "^[^@]+@[^@]+\.[^@]+$") {
        Write-Warning "Skipping user with invalid email format: $UserPrincipalName"
        continue
    }
    
    Write-Host "`n--- Processing user: $DisplayName ($UserPrincipalName) ---" -ForegroundColor Yellow
    Write-Host "This will export ALL chat history for this user" -ForegroundColor Yellow
    
    # Create safe search name
    $SafeUserName = $UserPrincipalName -replace '[^a-zA-Z0-9@._-]', '_'
    $SafeUserName = $SafeUserName -replace '@', '_AT_' -replace '\.', '_DOT_'
    $SearchName = "ChatHistoryExport_$($SafeUserName)_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $ExportName = "Export_$SearchName"
    
    try {
        # Create content search for all chat history
        $SearchCreated = New-ContentSearch -SearchName $SearchName -UserPrincipalName $UserPrincipalName
        
        if ($SearchCreated) {
            # Start and monitor search
            $SearchCompleted = Start-AndMonitorSearch -SearchName $SearchName
            
            if ($SearchCompleted) {
                # Export results
                $ExportInfo = Export-SearchResults -SearchName $SearchName -ExportName $ExportName -OutputPath $OutputDirectory
                
                $UserResult = [PSCustomObject]@{
                    UserPrincipalName = $UserPrincipalName
                    DisplayName = $DisplayName
                    SearchName = $SearchName
                    ExportName = $ExportName
                    Status = if ($ExportInfo -and $ExportInfo.Status -eq "Completed") { "Success" } elseif ($ExportInfo) { "Export $($ExportInfo.Status)" } else { "Export Failed" }
                    ExportDetails = $ExportInfo
                    ProcessedAt = Get-Date
                    ContentTypes = $IncludedTypes -join ', '
                }
            }
            else {
                $UserResult = [PSCustomObject]@{
                    UserPrincipalName = $UserPrincipalName
                    DisplayName = $DisplayName
                    SearchName = $SearchName
                    ExportName = $ExportName
                    Status = "Search Failed"
                    ExportDetails = $null
                    ProcessedAt = Get-Date
                    ContentTypes = $IncludedTypes -join ', '
                }
            }
        }
        else {
            $UserResult = [PSCustomObject]@{
                UserPrincipalName = $UserPrincipalName
                DisplayName = $DisplayName
                SearchName = $SearchName
                ExportName = $ExportName
                Status = "Search Creation Failed"
                ExportDetails = $null
                ProcessedAt = Get-Date
                ContentTypes = $IncludedTypes -join ', '
            }
        }
    }
    catch {
        Write-Error "Error processing user $UserPrincipalName`: $($_.Exception.Message)"
        $UserResult = [PSCustomObject]@{
            UserPrincipalName = $UserPrincipalName
            DisplayName = $DisplayName
            SearchName = $SearchName
            ExportName = $ExportName
            Status = "Error: $($_.Exception.Message)"
            ExportDetails = $null
            ProcessedAt = Get-Date
            ContentTypes = $IncludedTypes -join ', '
        }
    }
    
    $Results += $UserResult
    
    # Add a small delay between users to avoid overwhelming the service
    if ($Users.IndexOf($User) -lt ($Users.Count - 1)) {
        Write-Host "Waiting 30 seconds before processing next user..." -ForegroundColor Gray
        Start-Sleep -Seconds 30
    }
}

# Export results summary
$ResultsPath = Join-Path $OutputDirectory "ChatHistoryExportResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$Results | Export-Csv -Path $ResultsPath -NoTypeInformation

Write-Host "`n=== Export Summary ===" -ForegroundColor Cyan
Write-Host "Total users processed: $($Results.Count)" -ForegroundColor Green
Write-Host "Successful exports: $(($Results | Where-Object {$_.Status -eq 'Success'}).Count)" -ForegroundColor Green
Write-Host "Failed/Pending exports: $(($Results | Where-Object {$_.Status -ne 'Success'}).Count)" -ForegroundColor Red
Write-Host "Results saved to: $ResultsPath" -ForegroundColor Gray

# Display export instructions
Write-Host "`n=== Download Instructions ===" -ForegroundColor Cyan
Write-Host "1. Go to Security & Compliance Center (https://compliance.microsoft.com)" -ForegroundColor White
Write-Host "2. Navigate to Content Search > Export tab" -ForegroundColor White
Write-Host "3. Click on each completed export to download the data" -ForegroundColor White
Write-Host "4. Use the Microsoft Office 365 eDiscovery Export Tool for bulk downloads" -ForegroundColor White
Write-Host "5. Large chat histories may take several hours to complete" -ForegroundColor Yellow

Write-Host "`nScript completed at: $(Get-Date)" -ForegroundColor Gray
Write-Host "Note: Some exports may still be processing. Check the Compliance Center for final status." -ForegroundColor Yellow