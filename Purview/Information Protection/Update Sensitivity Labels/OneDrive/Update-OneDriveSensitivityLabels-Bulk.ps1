<#
.SYNOPSIS
    Scans multiple OneDrive accounts from a CSV file for files with a specific sensitivity label 
    and updates them to a new label using the Microsoft Graph API.

.DESCRIPTION
    This script reads a list of OneDrive users from a CSV file, connects to each user's OneDrive,
    scans for files with the source sensitivity label GUID, updates matching files to the target label GUID,
    and logs all changes to a CSV file.

.PARAMETER UsersCSVPath
    Path to the CSV file containing OneDrive user information.
    Required columns: UserPrincipalName or UserId

.PARAMETER SourceLabelId
    The GUID of the sensitivity label to search for. (Required)

.PARAMETER TargetLabelId
    The GUID of the sensitivity label to apply to matching files. (Required)

.PARAMETER LogPath
    The path for the CSV log file (default: current directory with timestamp).

.PARAMETER SummaryLogPath
    The path for the summary CSV log file showing per-user results.

.PARAMETER ClientId
    The Azure AD Application (Client) ID for authentication.

.PARAMETER TenantId
    The Azure AD Tenant ID.

.PARAMETER ClientSecret
    The Azure AD Application Client Secret.

.PARAMETER MaxConcurrentUsers
    Maximum number of users to process (for testing). Default processes all.

.PARAMETER ContinueOnError
    If specified, continues processing other users if one fails.

.EXAMPLE
    .\Update-OneDriveSensitivityLabels-Bulk.ps1 -UsersCSVPath ".\users.csv" `
        -SourceLabelId "your-source-label-guid" -TargetLabelId "your-target-label-guid" `
        -ClientId "your-client-id" -TenantId "your-tenant-id" -ClientSecret "your-secret"

.EXAMPLE
    # CSV file format (users.csv):
    # UserPrincipalName
    # john.doe@contoso.com
    # jane.smith@contoso.com
    # bob.wilson@contoso.com

.NOTES
    Prerequisites:
    - Azure AD App Registration with metered API billing enabled
    - Permissions: Files.ReadWrite.All, User.Read.All, InformationProtectionPolicy.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$UsersCSVPath,

    [Parameter(Mandatory = $true)]
    [string]$SourceLabelId,

    [Parameter(Mandatory = $true)]
    [string]$TargetLabelId,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\OneDriveLabelChanges_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter(Mandatory = $false)]
    [string]$SummaryLogPath = ".\OneDriveLabelSummary_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,

    [Parameter(Mandatory = $false)]
    [int]$MaxConcurrentUsers = 0,

    [Parameter(Mandatory = $false)]
    [switch]$ContinueOnError
)

#region Logging Functions
function Initialize-LogFile {
    param([string]$Path)
    
    $logDir = Split-Path -Path $Path -Parent
    if ($logDir -and -not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    $header = @(
        "Timestamp",
        "UserPrincipalName",
        "UserId",
        "FileName",
        "FileUrl",
        "FilePath",
        "OriginalLabelId",
        "OriginalLabelName",
        "NewLabelId",
        "NewLabelName",
        "Status",
        "ErrorMessage",
        "FileSize",
        "LastModified"
    )
    
    $header -join "," | Out-File -FilePath $Path -Encoding UTF8
    return $Path
}

function Initialize-SummaryLog {
    param([string]$Path)
    
    $logDir = Split-Path -Path $Path -Parent
    if ($logDir -and -not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    $header = @(
        "Timestamp",
        "UserPrincipalName",
        "UserId",
        "TotalFiles",
        "FilesScanned",
        "FilesMatched",
        "FilesUpdated",
        "FilesFailed",
        "Status",
        "ErrorMessage",
        "ProcessingTime"
    )
    
    $header -join "," | Out-File -FilePath $Path -Encoding UTF8
    return $Path
}

function Write-LogEntry {
    param(
        [string]$LogPath,
        [hashtable]$Entry
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    $logEntry = [PSCustomObject]@{
        Timestamp         = $timestamp
        UserPrincipalName = $Entry.UserPrincipalName ?? ""
        UserId            = $Entry.UserId ?? ""
        FileName          = $Entry.FileName ?? ""
        FileUrl           = $Entry.FileUrl ?? ""
        FilePath          = $Entry.FilePath ?? ""
        OriginalLabelId   = $Entry.OriginalLabelId ?? ""
        OriginalLabelName = $Entry.OriginalLabelName ?? ""
        NewLabelId        = $Entry.NewLabelId ?? ""
        NewLabelName      = $Entry.NewLabelName ?? ""
        Status            = $Entry.Status ?? ""
        ErrorMessage      = ($Entry.ErrorMessage ?? "") -replace "`n|`r", " "
        FileSize          = $Entry.FileSize ?? ""
        LastModified      = $Entry.LastModified ?? ""
    }
    
    $logEntry | Export-Csv -Path $LogPath -Append -NoTypeInformation -Encoding UTF8
}

function Write-SummaryEntry {
    param(
        [string]$LogPath,
        [hashtable]$Entry
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    $summaryEntry = [PSCustomObject]@{
        Timestamp         = $timestamp
        UserPrincipalName = $Entry.UserPrincipalName ?? ""
        UserId            = $Entry.UserId ?? ""
        TotalFiles        = $Entry.TotalFiles ?? 0
        FilesScanned      = $Entry.FilesScanned ?? 0
        FilesMatched      = $Entry.FilesMatched ?? 0
        FilesUpdated      = $Entry.FilesUpdated ?? 0
        FilesFailed       = $Entry.FilesFailed ?? 0
        Status            = $Entry.Status ?? ""
        ErrorMessage      = ($Entry.ErrorMessage ?? "") -replace "`n|`r", " "
        ProcessingTime    = $Entry.ProcessingTime ?? ""
    }
    
    $summaryEntry | Export-Csv -Path $LogPath -Append -NoTypeInformation -Encoding UTF8
}
#endregion

#region Authentication
function Get-GraphAccessToken {
    param(
        [string]$ClientId,
        [string]$TenantId,
        [string]$ClientSecret
    )
    
    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }
    
    try {
        $response = Invoke-RestMethod -Uri $tokenEndpoint -Method POST -Body $body -ContentType "application/x-www-form-urlencoded"
        return $response.access_token
    }
    catch {
        Write-Error "Failed to acquire Graph access token: $_"
        throw
    }
}
#endregion

#region Graph API Operations
function Get-UserDriveId {
    param(
        [string]$UserPrincipalName,
        [string]$AccessToken
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }
    
    try {
        # Get the user's OneDrive
        $uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/drive"
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET
        
        return @{
            DriveId = $response.id
            UserId  = $response.owner.user.id
            WebUrl  = $response.webUrl
        }
    }
    catch {
        throw "Failed to get OneDrive for user $UserPrincipalName : $_"
    }
}

function Get-AllOneDriveFiles {
    param(
        [string]$UserId,
        [string]$DriveId,
        [string]$AccessToken,
        [string]$FolderId = ""
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }
    
    $allFiles = @()
    
    try {
        # Use item ID-based navigation instead of path-based
        if ($FolderId) {
            $uri = "https://graph.microsoft.com/v1.0/users/$UserId/drive/items/$FolderId/children"
        }
        else {
            $uri = "https://graph.microsoft.com/v1.0/users/$UserId/drive/root/children"
        }
        
        do {
            try {
                $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET
            }
            catch {
                # If folder is empty or inaccessible, return empty array
                if ($_.Exception.Response.StatusCode -eq 404) {
                    Write-Verbose "Folder not found or empty: $FolderId"
                    return $allFiles
                }
                throw
            }
            
            if ($response.value) {
                foreach ($item in $response.value) {
                    if ($item.folder) {
                        # Recursively get files from subfolders using item ID
                        try {
                            $subFiles = Get-AllOneDriveFiles -UserId $UserId -DriveId $DriveId -AccessToken $AccessToken -FolderId $item.id
                            $allFiles += $subFiles
                        }
                        catch {
                            Write-Verbose "Could not access folder '$($item.name)': $_"
                            # Continue with other folders
                        }
                    }
                    else {
                        # It's a file
                        $allFiles += $item
                    }
                }
            }
            
            $uri = $response.'@odata.nextLink'
        } while ($uri)
        
        return $allFiles
    }
    catch {
        throw "Failed to enumerate OneDrive files: $_"
    }
}

function Get-FileSensitivityLabelViaGraph {
    param(
        [string]$UserId,
        [string]$DriveId,
        [string]$ItemId,
        [string]$AccessToken
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/users/$UserId/drive/items/$ItemId/extractSensitivityLabels"
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method POST
        
        if ($response.labels -and $response.labels.Count -gt 0) {
            return $response.labels[0]
        }
        return $null
    }
    catch {
        if ($_.Exception.Response.StatusCode -ne 404) {
            Write-Verbose "Error getting sensitivity label for item $ItemId : $_"
        }
        return $null
    }
}

function Set-FileSensitivityLabelViaGraph {
    param(
        [string]$UserId,
        [string]$DriveId,
        [string]$ItemId,
        [string]$LabelId,
        [string]$AccessToken,
        [string]$JustificationMessage = "Automated label update via PowerShell script"
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }
    
    $body = @{
        sensitivityLabelId = $LabelId
        assignmentMethod   = "auto"
        justificationText  = $JustificationMessage
    } | ConvertTo-Json
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/users/$UserId/drive/items/$ItemId/assignSensitivityLabel"
        Invoke-RestMethod -Uri $uri -Headers $headers -Method POST -Body $body | Out-Null
        return $true
    }
    catch {
        throw "Error setting sensitivity label: $_"
    }
}

function Get-FileLabelInfo {
    param(
        [object]$File,
        [string]$UserId,
        [string]$DriveId,
        [string]$AccessToken
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/users/$UserId/drive/items/$($File.id)?`$select=id,name,sensitivityLabel"
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET
        
        if ($response.sensitivityLabel) {
            return @{
                LabelId   = $response.sensitivityLabel.id
                LabelName = $response.sensitivityLabel.displayName
            }
        }
        
        $labelInfo = Get-FileSensitivityLabelViaGraph -UserId $UserId -DriveId $DriveId -ItemId $File.id -AccessToken $AccessToken
        if ($labelInfo) {
            return @{
                LabelId   = $labelInfo.sensitivityLabelId
                LabelName = $labelInfo.name ?? "Unknown"
            }
        }
        
        return $null
    }
    catch {
        Write-Verbose "Could not get label info for $($File.name): $_"
        return $null
    }
}
#endregion

#region User Processing
function Process-SingleUser {
    param(
        [string]$UserPrincipalName,
        [string]$SourceLabelId,
        [string]$TargetLabelId,
        [string]$AccessToken,
        [string]$LogPath
    )
    
    $userStats = @{
        TotalFiles   = 0
        FilesScanned = 0
        FilesMatched = 0
        FilesUpdated = 0
        FilesFailed  = 0
        UserId       = ""
    }
    
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    try {
        # Get user's OneDrive info
        $driveInfo = Get-UserDriveId -UserPrincipalName $UserPrincipalName -AccessToken $AccessToken
        $userStats.UserId = $driveInfo.UserId
        Write-Host "  OneDrive found: $($driveInfo.WebUrl)" -ForegroundColor Gray
        
        # Get all files (no FolderId = start from root)
        $files = Get-AllOneDriveFiles -UserId $UserPrincipalName -DriveId $driveInfo.DriveId -AccessToken $AccessToken
        $userStats.TotalFiles = $files.Count
        Write-Host "  Found $($files.Count) files" -ForegroundColor Gray
        
        # Process each file
        foreach ($file in $files) {
            $userStats.FilesScanned++
            
            $labelInfo = Get-FileLabelInfo -File $file -UserId $UserPrincipalName -DriveId $driveInfo.DriveId -AccessToken $AccessToken
            
            if (-not $labelInfo) { continue }
            
            if ($labelInfo.LabelId -eq $SourceLabelId) {
                $userStats.FilesMatched++
                
                $logEntry = @{
                    UserPrincipalName = $UserPrincipalName
                    UserId            = $driveInfo.UserId
                    FileName          = $file.name
                    FileUrl           = $file.webUrl
                    FilePath          = $file.parentReference.path
                    OriginalLabelId   = $labelInfo.LabelId
                    OriginalLabelName = $labelInfo.LabelName
                    NewLabelId        = $TargetLabelId
                    FileSize          = $file.size
                    LastModified      = $file.lastModifiedDateTime
                }
                
                try {
                    $updateResult = Set-FileSensitivityLabelViaGraph `
                        -UserId $UserPrincipalName `
                        -DriveId $driveInfo.DriveId `
                        -ItemId $file.id `
                        -LabelId $TargetLabelId `
                        -AccessToken $AccessToken
                    
                    $userStats.FilesUpdated++
                    $logEntry.Status = "Updated"
                    $logEntry.NewLabelName = "Label Applied"
                    Write-Host "    Updated: $($file.name)" -ForegroundColor Green
                }
                catch {
                    $userStats.FilesFailed++
                    $logEntry.Status = "Failed"
                    $logEntry.ErrorMessage = $_.Exception.Message
                    Write-Host "    Failed: $($file.name) - $($_.Exception.Message)" -ForegroundColor Red
                }
                
                Write-LogEntry -LogPath $LogPath -Entry $logEntry
            }
        }
        
        $stopwatch.Stop()
        $userStats.ProcessingTime = $stopwatch.Elapsed.ToString("hh\:mm\:ss")
        $userStats.Status = "Completed"
        
        return $userStats
    }
    catch {
        $stopwatch.Stop()
        $userStats.ProcessingTime = $stopwatch.Elapsed.ToString("hh\:mm\:ss")
        $userStats.Status = "Error"
        $userStats.ErrorMessage = $_.Exception.Message
        throw
    }
}
#endregion

#region Main Processing
function Start-BulkLabelMigration {
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "OneDrive Bulk Sensitivity Label Migration" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Read CSV file
    Write-Host "Reading users from: $UsersCSVPath" -ForegroundColor Yellow
    $users = Import-Csv -Path $UsersCSVPath
    
    # Validate CSV structure
    $hasUPN = $users[0].PSObject.Properties.Name.Contains('UserPrincipalName')
    $hasEmail = $users[0].PSObject.Properties.Name.Contains('Email')
    $hasUserId = $users[0].PSObject.Properties.Name.Contains('UserId')
    
    if (-not ($hasUPN -or $hasEmail -or $hasUserId)) {
        throw "CSV file must contain a 'UserPrincipalName', 'Email', or 'UserId' column"
    }
    
    $totalUsers = $users.Count
    if ($MaxConcurrentUsers -gt 0 -and $MaxConcurrentUsers -lt $totalUsers) {
        $users = $users | Select-Object -First $MaxConcurrentUsers
        Write-Host "Processing first $MaxConcurrentUsers users (of $totalUsers total)" -ForegroundColor Yellow
    }
    else {
        Write-Host "Processing $totalUsers users" -ForegroundColor Yellow
    }
    
    Write-Host "Source Label: $SourceLabelId" -ForegroundColor Yellow
    Write-Host "Target Label: $TargetLabelId" -ForegroundColor Green
    Write-Host ""
    
    # Initialize log files
    $detailLogFile = Initialize-LogFile -Path $LogPath
    $summaryLogFile = Initialize-SummaryLog -Path $SummaryLogPath
    Write-Host "Detail log: $detailLogFile" -ForegroundColor Gray
    Write-Host "Summary log: $summaryLogFile" -ForegroundColor Gray
    Write-Host ""
    
    # Global statistics
    $globalStats = @{
        TotalUsers       = $users.Count
        UsersProcessed   = 0
        UsersSucceeded   = 0
        UsersFailed      = 0
        TotalFiles       = 0
        TotalMatched     = 0
        TotalUpdated     = 0
        TotalFailed      = 0
    }
    
    $overallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    try {
        # Get access token
        Write-Host "Acquiring access token..." -ForegroundColor Yellow
        $accessToken = Get-GraphAccessToken -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret
        Write-Host "Access token acquired" -ForegroundColor Green
        Write-Host ""
        
        # Process each user
        $userNumber = 0
        foreach ($user in $users) {
            $userNumber++
            
            # Get user identifier from CSV (support multiple column names)
            $userPrincipalName = if ($user.UserPrincipalName) { $user.UserPrincipalName.Trim() }
                                 elseif ($user.Email) { $user.Email.Trim() }
                                 elseif ($user.UserId) { $user.UserId.Trim() }
                                 else { $null }
            
            if (-not $userPrincipalName) {
                Write-Host "[$userNumber/$($users.Count)] Skipping row - no user identifier found" -ForegroundColor Yellow
                continue
            }
            
            Write-Host "[$userNumber/$($users.Count)] Processing: $userPrincipalName" -ForegroundColor Cyan
            
            $summaryEntry = @{
                UserPrincipalName = $userPrincipalName
            }
            
            try {
                $userStats = Process-SingleUser `
                    -UserPrincipalName $userPrincipalName `
                    -SourceLabelId $SourceLabelId `
                    -TargetLabelId $TargetLabelId `
                    -AccessToken $accessToken `
                    -LogPath $detailLogFile
                
                $globalStats.UsersSucceeded++
                $globalStats.TotalFiles += $userStats.TotalFiles
                $globalStats.TotalMatched += $userStats.FilesMatched
                $globalStats.TotalUpdated += $userStats.FilesUpdated
                $globalStats.TotalFailed += $userStats.FilesFailed
                
                $summaryEntry.UserId = $userStats.UserId
                $summaryEntry.TotalFiles = $userStats.TotalFiles
                $summaryEntry.FilesScanned = $userStats.FilesScanned
                $summaryEntry.FilesMatched = $userStats.FilesMatched
                $summaryEntry.FilesUpdated = $userStats.FilesUpdated
                $summaryEntry.FilesFailed = $userStats.FilesFailed
                $summaryEntry.Status = "Completed"
                $summaryEntry.ProcessingTime = $userStats.ProcessingTime
                
                Write-Host "  Completed: $($userStats.FilesUpdated) updated, $($userStats.FilesFailed) failed" -ForegroundColor Green
            }
            catch {
                $globalStats.UsersFailed++
                $summaryEntry.Status = "Error"
                $summaryEntry.ErrorMessage = $_.Exception.Message
                
                Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
                
                if (-not $ContinueOnError) {
                    throw "User processing failed and -ContinueOnError not specified"
                }
            }
            finally {
                $globalStats.UsersProcessed++
                Write-SummaryEntry -LogPath $summaryLogFile -Entry $summaryEntry
            }
            
            Write-Host ""
            
            # Refresh token if needed (tokens typically last 1 hour)
            if ($userNumber % 10 -eq 0) {
                Write-Host "Refreshing access token..." -ForegroundColor Yellow
                $accessToken = Get-GraphAccessToken -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret
            }
        }
        
        $overallStopwatch.Stop()
    }
    catch {
        $overallStopwatch.Stop()
        Write-Error "Bulk migration failed: $_"
    }
    finally {
        # Display final summary
        Write-Host "==========================================" -ForegroundColor Cyan
        Write-Host "Migration Complete - Final Summary" -ForegroundColor Cyan
        Write-Host "==========================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Users:" -ForegroundColor White
        Write-Host "  Total users: $($globalStats.TotalUsers)" -ForegroundColor White
        Write-Host "  Succeeded: $($globalStats.UsersSucceeded)" -ForegroundColor Green
        Write-Host "  Failed: $($globalStats.UsersFailed)" -ForegroundColor $(if ($globalStats.UsersFailed -gt 0) { "Red" } else { "White" })
        Write-Host ""
        Write-Host "Files:" -ForegroundColor White
        Write-Host "  Total files scanned: $($globalStats.TotalFiles)" -ForegroundColor White
        Write-Host "  Files matching source label: $($globalStats.TotalMatched)" -ForegroundColor Yellow
        Write-Host "  Successfully updated: $($globalStats.TotalUpdated)" -ForegroundColor Green
        Write-Host "  Failed to update: $($globalStats.TotalFailed)" -ForegroundColor $(if ($globalStats.TotalFailed -gt 0) { "Red" } else { "White" })
        Write-Host ""
        Write-Host "Total processing time: $($overallStopwatch.Elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor White
        Write-Host ""
        Write-Host "Log files:" -ForegroundColor White
        Write-Host "  Detail log: $detailLogFile" -ForegroundColor Cyan
        Write-Host "  Summary log: $summaryLogFile" -ForegroundColor Cyan
        Write-Host ""
    }
    
    return $globalStats
}
#endregion

#region Script Entry Point
try {
    $results = Start-BulkLabelMigration
    
    if ($results.UsersFailed -gt 0 -or $results.TotalFailed -gt 0) {
        exit 1
    }
    exit 0
}
catch {
    Write-Error "Script execution failed: $_"
    Write-Error $_.ScriptStackTrace
    exit 1
}
#endregion
