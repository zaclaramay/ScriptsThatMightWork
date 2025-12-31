<#
.SYNOPSIS
    Scans multiple SharePoint Online sites from a CSV file for files with a specific sensitivity label 
    and updates them to a new label using the Microsoft Graph API.

.DESCRIPTION
    This script reads a list of SharePoint sites from a CSV file, connects to each site's document library,
    scans for files with the source sensitivity label GUID, updates matching files to the target label GUID,
    and logs all changes to a CSV file.

.PARAMETER SitesCSVPath
    Path to the CSV file containing SharePoint site information.
    Required columns: SiteUrl, LibraryName (optional, defaults to "Documents")

.PARAMETER SourceLabelId
    The GUID of the sensitivity label to search for.

.PARAMETER TargetLabelId
    The GUID of the sensitivity label to apply to matching files.

.PARAMETER LogPath
    The path for the CSV log file (default: current directory with timestamp).

.PARAMETER SummaryLogPath
    The path for the summary CSV log file showing per-site results.

.PARAMETER ClientId
    The Azure AD Application (Client) ID for authentication.

.PARAMETER TenantId
    The Azure AD Tenant ID.

.PARAMETER ClientSecret
    The Azure AD Application Client Secret.

.PARAMETER MaxConcurrentSites
    Maximum number of sites to process (for testing). Default processes all.

.PARAMETER ContinueOnError
    If specified, continues processing other sites if one fails.

.EXAMPLE
    .\Update-SharePointSensitivityLabels-Bulk.ps1 -SitesCSVPath ".\sites.csv" `
        -ClientId "your-client-id" -TenantId "your-tenant-id" -ClientSecret "your-secret"

.EXAMPLE
    # CSV file format (sites.csv):
    # SiteUrl,LibraryName
    # https://contoso.sharepoint.com/sites/IT,Documents
    # https://contoso.sharepoint.com/sites/HR,Shared Documents
    # https://contoso.sharepoint.com/sites/Finance,Documents

.NOTES
    Author: Claude Assistant
    Prerequisites:
    - Azure AD App Registration with metered API billing enabled
    - Permissions: Sites.ReadWrite.All, Files.ReadWrite.All, InformationProtectionPolicy.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$SitesCSVPath,

    [Parameter(Mandatory = $false)]
    [string]$SourceLabelId = "", # set your default target source label GUID here

    [Parameter(Mandatory = $false)]
    [string]$TargetLabelId = "", # set your default target label GUID here

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\SensitivityLabelChanges_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv", # path for detailed log

    [Parameter(Mandatory = $false)]
    [string]$SummaryLogPath = ".\SensitivityLabelSummary_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv", #path for summary log

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,

    [Parameter(Mandatory = $false)]
    [int]$MaxConcurrentSites = 0,

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
        "SiteUrl",
        "LibraryName",
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
        "ModifiedBy",
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
        "SiteUrl",
        "LibraryName",
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
        SiteUrl           = $Entry.SiteUrl ?? ""
        LibraryName       = $Entry.LibraryName ?? ""
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
        ModifiedBy        = $Entry.ModifiedBy ?? ""
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
        Timestamp      = $timestamp
        SiteUrl        = $Entry.SiteUrl ?? ""
        LibraryName    = $Entry.LibraryName ?? ""
        TotalFiles     = $Entry.TotalFiles ?? 0
        FilesScanned   = $Entry.FilesScanned ?? 0
        FilesMatched   = $Entry.FilesMatched ?? 0
        FilesUpdated   = $Entry.FilesUpdated ?? 0
        FilesFailed    = $Entry.FilesFailed ?? 0
        Status         = $Entry.Status ?? ""
        ErrorMessage   = ($Entry.ErrorMessage ?? "") -replace "`n|`r", " "
        ProcessingTime = $Entry.ProcessingTime ?? ""
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
function Get-SiteAndDriveIds {
    param(
        [string]$SiteUrl,
        [string]$LibraryName,
        [string]$AccessToken
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }
    
    $uri = [System.Uri]$SiteUrl
    $hostname = $uri.Host
    $sitePath = $uri.AbsolutePath.TrimEnd('/')
    
    try {
        # Get site ID
        $siteUri = "https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}"
        $siteResponse = Invoke-RestMethod -Uri $siteUri -Headers $headers -Method GET
        $siteId = $siteResponse.id
        
        # Get drive ID for the library
        $drivesUri = "https://graph.microsoft.com/v1.0/sites/$siteId/drives"
        $drivesResponse = Invoke-RestMethod -Uri $drivesUri -Headers $headers -Method GET
        
        $drive = $drivesResponse.value | Where-Object { $_.name -eq $LibraryName }
        
        if (-not $drive) {
            # Try matching by web URL or other properties
            $drive = $drivesResponse.value | Where-Object { 
                $_.name -like "*$LibraryName*" -or $_.webUrl -like "*$LibraryName*" 
            } | Select-Object -First 1
        }
        
        if (-not $drive) {
            throw "Document library '$LibraryName' not found. Available libraries: $($drivesResponse.value.name -join ', ')"
        }
        
        return @{
            SiteId    = $siteId
            DriveId   = $drive.id
            DriveName = $drive.name
        }
    }
    catch {
        throw "Failed to get site/drive IDs for $SiteUrl : $_"
    }
}

function Get-AllLibraryFiles {
    param(
        [string]$SiteId,
        [string]$DriveId,
        [string]$AccessToken,
        [string]$FolderPath = ""
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }
    
    $allFiles = @()
    
    try {
        if ($FolderPath) {
            $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/root:/${FolderPath}:/children"
        }
        else {
            $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/root/children"
        }
        
        do {
            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET
            
            foreach ($item in $response.value) {
                if ($item.folder) {
                    $subPath = if ($FolderPath) { "$FolderPath/$($item.name)" } else { $item.name }
                    $subFiles = Get-AllLibraryFiles -SiteId $SiteId -DriveId $DriveId -AccessToken $AccessToken -FolderPath $subPath
                    $allFiles += $subFiles
                }
                else {
                    $allFiles += $item
                }
            }
            
            $uri = $response.'@odata.nextLink'
        } while ($uri)
        
        return $allFiles
    }
    catch {
        throw "Failed to enumerate library files: $_"
    }
}

function Get-FileSensitivityLabelViaGraph {
    param(
        [string]$SiteId,
        [string]$DriveId,
        [string]$ItemId,
        [string]$AccessToken
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$ItemId/extractSensitivityLabels"
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
        [string]$SiteId,
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
        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$ItemId/assignSensitivityLabel"
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
        [string]$SiteId,
        [string]$DriveId,
        [string]$AccessToken
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$($File.id)?`$select=id,name,sensitivityLabel"
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET
        
        if ($response.sensitivityLabel) {
            return @{
                LabelId   = $response.sensitivityLabel.id
                LabelName = $response.sensitivityLabel.displayName
            }
        }
        
        $labelInfo = Get-FileSensitivityLabelViaGraph -SiteId $SiteId -DriveId $DriveId -ItemId $File.id -AccessToken $AccessToken
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

#region Site Processing
function Process-SingleSite {
    param(
        [string]$SiteUrl,
        [string]$LibraryName,
        [string]$SourceLabelId,
        [string]$TargetLabelId,
        [string]$AccessToken,
        [string]$LogPath
    )
    
    $siteStats = @{
        TotalFiles   = 0
        FilesScanned = 0
        FilesMatched = 0
        FilesUpdated = 0
        FilesFailed  = 0
    }
    
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    try {
        # Get site and drive info
        $siteInfo = Get-SiteAndDriveIds -SiteUrl $SiteUrl -LibraryName $LibraryName -AccessToken $AccessToken
        Write-Host "  Library found: $($siteInfo.DriveName)" -ForegroundColor Gray
        
        # Get all files
        $files = Get-AllLibraryFiles -SiteId $siteInfo.SiteId -DriveId $siteInfo.DriveId -AccessToken $AccessToken
        $siteStats.TotalFiles = $files.Count
        Write-Host "  Found $($files.Count) files" -ForegroundColor Gray
        
        # Process each file
        foreach ($file in $files) {
            $siteStats.FilesScanned++
            
            $labelInfo = Get-FileLabelInfo -File $file -SiteId $siteInfo.SiteId -DriveId $siteInfo.DriveId -AccessToken $AccessToken
            
            if (-not $labelInfo) { continue }
            
            if ($labelInfo.LabelId -eq $SourceLabelId) {
                $siteStats.FilesMatched++
                
                $logEntry = @{
                    SiteUrl           = $SiteUrl
                    LibraryName       = $LibraryName
                    FileName          = $file.name
                    FileUrl           = $file.webUrl
                    FilePath          = $file.parentReference.path
                    OriginalLabelId   = $labelInfo.LabelId
                    OriginalLabelName = $labelInfo.LabelName
                    NewLabelId        = $TargetLabelId
                    FileSize          = $file.size
                    ModifiedBy        = $file.lastModifiedBy.user.displayName
                    LastModified      = $file.lastModifiedDateTime
                }
                
                try {
                    $updateResult = Set-FileSensitivityLabelViaGraph `
                        -SiteId $siteInfo.SiteId `
                        -DriveId $siteInfo.DriveId `
                        -ItemId $file.id `
                        -LabelId $TargetLabelId `
                        -AccessToken $AccessToken
                    
                    $siteStats.FilesUpdated++
                    $logEntry.Status = "Updated"
                    $logEntry.NewLabelName = "Label Applied"
                    Write-Host "    Updated: $($file.name)" -ForegroundColor Green
                }
                catch {
                    $siteStats.FilesFailed++
                    $logEntry.Status = "Failed"
                    $logEntry.ErrorMessage = $_.Exception.Message
                    Write-Host "    Failed: $($file.name) - $($_.Exception.Message)" -ForegroundColor Red
                }
                
                Write-LogEntry -LogPath $LogPath -Entry $logEntry
            }
        }
        
        $stopwatch.Stop()
        $siteStats.ProcessingTime = $stopwatch.Elapsed.ToString("hh\:mm\:ss")
        $siteStats.Status = "Completed"
        
        return $siteStats
    }
    catch {
        $stopwatch.Stop()
        $siteStats.ProcessingTime = $stopwatch.Elapsed.ToString("hh\:mm\:ss")
        $siteStats.Status = "Error"
        $siteStats.ErrorMessage = $_.Exception.Message
        throw
    }
}
#endregion

#region Main Processing
function Start-BulkLabelMigration {
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "SharePoint Bulk Sensitivity Label Migration" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Read CSV file
    Write-Host "Reading sites from: $SitesCSVPath" -ForegroundColor Yellow
    $sites = Import-Csv -Path $SitesCSVPath
    
    # Validate CSV structure
    if (-not $sites[0].PSObject.Properties.Name.Contains('SiteUrl')) {
        throw "CSV file must contain a 'SiteUrl' column"
    }
    
    $totalSites = $sites.Count
    if ($MaxConcurrentSites -gt 0 -and $MaxConcurrentSites -lt $totalSites) {
        $sites = $sites | Select-Object -First $MaxConcurrentSites
        Write-Host "Processing first $MaxConcurrentSites sites (of $totalSites total)" -ForegroundColor Yellow
    }
    else {
        Write-Host "Processing $totalSites sites" -ForegroundColor Yellow
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
        TotalSites       = $sites.Count
        SitesProcessed   = 0
        SitesSucceeded   = 0
        SitesFailed      = 0
        TotalFiles       = 0
        TotalMatched     = 0
        TotalUpdated     = 0
        TotalFailed      = 0
    }
    
    $overallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $resultsToReturn = $null
    
    try {
        # Get access token
        Write-Host "Acquiring access token..." -ForegroundColor Yellow
        $accessToken = Get-GraphAccessToken -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret
        Write-Host "Access token acquired" -ForegroundColor Green
        Write-Host ""
        
        # Process each site
        $siteNumber = 0
        foreach ($site in $sites) {
            $siteNumber++
            $siteUrl = $site.SiteUrl.Trim()
            $libraryName = if ($site.LibraryName) { $site.LibraryName.Trim() } else { "Documents" }
            
            Write-Host "[$siteNumber/$($sites.Count)] Processing: $siteUrl" -ForegroundColor Cyan
            Write-Host "  Library: $libraryName" -ForegroundColor Gray
            
            $summaryEntry = @{
                SiteUrl     = $siteUrl
                LibraryName = $libraryName
            }
            
            try {
                $siteStats = Process-SingleSite `
                    -SiteUrl $siteUrl `
                    -LibraryName $libraryName `
                    -SourceLabelId $SourceLabelId `
                    -TargetLabelId $TargetLabelId `
                    -AccessToken $accessToken `
                    -LogPath $detailLogFile
                
                $globalStats.SitesSucceeded++
                $globalStats.TotalFiles += $siteStats.TotalFiles
                $globalStats.TotalMatched += $siteStats.FilesMatched
                $globalStats.TotalUpdated += $siteStats.FilesUpdated
                $globalStats.TotalFailed += $siteStats.FilesFailed
                
                $summaryEntry.TotalFiles = $siteStats.TotalFiles
                $summaryEntry.FilesScanned = $siteStats.FilesScanned
                $summaryEntry.FilesMatched = $siteStats.FilesMatched
                $summaryEntry.FilesUpdated = $siteStats.FilesUpdated
                $summaryEntry.FilesFailed = $siteStats.FilesFailed
                $summaryEntry.Status = "Completed"
                $summaryEntry.ProcessingTime = $siteStats.ProcessingTime
                
                Write-Host "  Completed: $($siteStats.FilesUpdated) updated, $($siteStats.FilesFailed) failed" -ForegroundColor Green
            }
            catch {
                $globalStats.SitesFailed++
                $summaryEntry.Status = "Error"
                $summaryEntry.ErrorMessage = $_.Exception.Message
                
                Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
                
                if (-not $ContinueOnError) {
                    throw "Site processing failed and -ContinueOnError not specified"
                }
            }
            finally {
                $globalStats.SitesProcessed++
                Write-SummaryEntry -LogPath $summaryLogFile -Entry $summaryEntry
            }
            
            Write-Host ""
            
            # Refresh token if needed (tokens typically last 1 hour)
            if ($siteNumber % 10 -eq 0) {
                Write-Host "Refreshing access token..." -ForegroundColor Yellow
                $accessToken = Get-GraphAccessToken -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret
            }
        }
        
        $overallStopwatch.Stop()
        $resultsToReturn = $globalStats
    }
    catch {
        $overallStopwatch.Stop()
        Write-Error "Bulk migration failed: $_"
        $resultsToReturn = $globalStats
    }
    finally {
        # Display final summary
        Write-Host "==========================================" -ForegroundColor Cyan
        Write-Host "Migration Complete - Final Summary" -ForegroundColor Cyan
        Write-Host "==========================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Sites:" -ForegroundColor White
        Write-Host "  Total sites: $($globalStats.TotalSites)" -ForegroundColor White
        Write-Host "  Succeeded: $($globalStats.SitesSucceeded)" -ForegroundColor Green
        Write-Host "  Failed: $($globalStats.SitesFailed)" -ForegroundColor $(if ($globalStats.SitesFailed -gt 0) { "Red" } else { "White" })
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
    
    return $resultsToReturn
}
#endregion

#region Script Entry Point
try {
    $results = Start-BulkLabelMigration
    
    if ($results -and ($results.SitesFailed -gt 0 -or $results.TotalFailed -gt 0)) {
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
