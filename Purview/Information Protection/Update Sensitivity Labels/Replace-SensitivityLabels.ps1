<#
.SYNOPSIS
    Scans a SharePoint Online document library for files with a specific sensitivity label 
    and updates them to a new label using the Microsoft Information Protection (MIP) SDK.

.DESCRIPTION
    This script connects to SharePoint Online, scans the specified document library for files
    with the source sensitivity label GUID, updates matching files to the target label GUID,
    and logs all changes to a CSV file.

.PARAMETER SiteUrl
    The SharePoint Online site URL containing the document library.

.PARAMETER LibraryName
    The name of the document library to scan (default: "Documents").

.PARAMETER SourceLabelId
    The GUID of the sensitivity label to search for.

.PARAMETER TargetLabelId
    The GUID of the sensitivity label to apply to matching files.

.PARAMETER LogPath
    The path for the CSV log file (default: current directory with timestamp).

.PARAMETER ClientId
    The Azure AD Application (Client) ID for authentication.

.PARAMETER TenantId
    The Azure AD Tenant ID.

.PARAMETER ClientSecret
    The Azure AD Application Client Secret (for app-only authentication).

.EXAMPLE
    .\Update-SharePointSensitivityLabels.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/InformationTechnology" `
        -ClientId "your-client-id" -TenantId "your-tenant-id" -ClientSecret "your-secret"

.NOTES
    Prerequisites:
    - PnP.PowerShell module
    - Microsoft.InformationProtection.File module (MIP SDK)
    - Azure AD App Registration with appropriate permissions:
      - Sites.ReadWrite.All (SharePoint)
      - InformationProtectionPolicy.Read (MIP)
    - Sensitivity labels must be published to users/groups
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$LibraryName = "", # Default document library name

    [Parameter(Mandatory = $false)]
    [string]$SourceLabelId = "", # The GUID of the source sensitivity label

    [Parameter(Mandatory = $false)]
    [string]$TargetLabelId = "", # The GUID of the target sensitivity label

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\SensitivityLabelChanges_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv", # Default log file path

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,

    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint
)

#region Module Verification and Installation
function Install-RequiredModules {
    $requiredModules = @(
        @{ Name = "PnP.PowerShell"; MinVersion = "2.0.0" }
    )

    foreach ($module in $requiredModules) {
        $installed = Get-Module -ListAvailable -Name $module.Name | 
            Where-Object { $_.Version -ge [version]$module.MinVersion }
        
        if (-not $installed) {
            Write-Host "Installing module: $($module.Name)..." -ForegroundColor Yellow
            try {
                Install-Module -Name $module.Name -MinimumVersion $module.MinVersion -Force -AllowClobber -Scope CurrentUser
                Write-Host "Successfully installed $($module.Name)" -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to install $($module.Name): $_"
                throw
            }
        }
        Import-Module -Name $module.Name -MinimumVersion $module.MinVersion -Force
    }
}
#endregion

#region Logging Functions
function Initialize-LogFile {
    param([string]$Path)
    
    $logDir = Split-Path -Path $Path -Parent
    if ($logDir -and -not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    # Create CSV header
    $header = @(
        "Timestamp",
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
    Write-Host "Log file initialized: $Path" -ForegroundColor Green
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
        FileName          = $Entry.FileName ?? ""
        FileUrl           = $Entry.FileUrl ?? ""
        FilePath          = $Entry.FilePath ?? ""
        OriginalLabelId   = $Entry.OriginalLabelId ?? ""
        OriginalLabelName = $Entry.OriginalLabelName ?? ""
        NewLabelId        = $Entry.NewLabelId ?? ""
        NewLabelName      = $Entry.NewLabelName ?? ""
        Status            = $Entry.Status ?? ""
        ErrorMessage      = $Entry.ErrorMessage ?? ""
        FileSize          = $Entry.FileSize ?? ""
        ModifiedBy        = $Entry.ModifiedBy ?? ""
        LastModified      = $Entry.LastModified ?? ""
    }
    
    $logEntry | Export-Csv -Path $LogPath -Append -NoTypeInformation -Encoding UTF8
}
#endregion

#region MIP SDK Integration
class AuthDelegateImplementation {
    [string]$ClientId
    [string]$TenantId
    [string]$ClientSecret
    [string]$AccessToken

    AuthDelegateImplementation([string]$clientId, [string]$tenantId, [string]$clientSecret) {
        $this.ClientId = $clientId
        $this.TenantId = $tenantId
        $this.ClientSecret = $clientSecret
    }

    [string] AcquireToken([string]$authority, [string]$resource, [string]$claims) {
        $tokenEndpoint = "https://login.microsoftonline.com/$($this.TenantId)/oauth2/v2.0/token"
        
        $body = @{
            client_id     = $this.ClientId
            client_secret = $this.ClientSecret
            scope         = "$resource/.default"
            grant_type    = "client_credentials"
        }
        
        try {
            $response = Invoke-RestMethod -Uri $tokenEndpoint -Method POST -Body $body -ContentType "application/x-www-form-urlencoded"
            $this.AccessToken = $response.access_token
            return $response.access_token
        }
        catch {
            Write-Error "Failed to acquire token: $_"
            throw
        }
    }
}

function Initialize-MipContext {
    param(
        [string]$ClientId,
        [string]$TenantId,
        [string]$ClientSecret
    )
    
    # Check for MIP SDK assemblies
    $mipAssemblyPath = $null
    $searchPaths = @(
        "$env:ProgramFiles\Microsoft Information Protection SDK\assemblies\net472",
        "$env:LOCALAPPDATA\Programs\MIP SDK\assemblies\net472",
        ".\MipSdk"
    )
    
    foreach ($path in $searchPaths) {
        if (Test-Path "$path\Microsoft.InformationProtection.File.dll") {
            $mipAssemblyPath = $path
            break
        }
    }
    
    if ($mipAssemblyPath) {
        try {
            Add-Type -Path "$mipAssemblyPath\Microsoft.InformationProtection.dll"
            Add-Type -Path "$mipAssemblyPath\Microsoft.InformationProtection.File.dll"
            Write-Host "MIP SDK loaded from: $mipAssemblyPath" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Warning "Failed to load MIP SDK assemblies: $_"
        }
    }
    
    Write-Warning "MIP SDK assemblies not found. Using Graph API fallback for label operations."
    return $false
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
        # Get file sensitivity label using Microsoft Graph
        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$ItemId/extractSensitivityLabels"
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method POST
        
        if ($response.labels -and $response.labels.Count -gt 0) {
            return $response.labels[0]
        }
        return $null
    }
    catch {
        if ($_.Exception.Response.StatusCode -ne 404) {
            Write-Warning "Error getting sensitivity label for item $ItemId : $_"
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
        sensitivityLabelId     = $LabelId
        assignmentMethod       = "auto"
        justificationText      = $JustificationMessage
    } | ConvertTo-Json
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$ItemId/assignSensitivityLabel"
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method POST -Body $body
        return $true
    }
    catch {
        Write-Error "Error setting sensitivity label for item $ItemId : $_"
        return $false
    }
}
#endregion

#region SharePoint Operations
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
    
    # Parse site URL to get hostname and site path
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
            throw "Document library '$LibraryName' not found in site"
        }
        
        return @{
            SiteId  = $siteId
            DriveId = $drive.id
        }
    }
    catch {
        Write-Error "Failed to get site/drive IDs: $_"
        throw
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
                    # Recursively get files from subfolders
                    $subPath = if ($FolderPath) { "$FolderPath/$($item.name)" } else { $item.name }
                    $subFiles = Get-AllLibraryFiles -SiteId $SiteId -DriveId $DriveId -AccessToken $AccessToken -FolderPath $subPath
                    $allFiles += $subFiles
                }
                else {
                    # It's a file
                    $allFiles += $item
                }
            }
            
            $uri = $response.'@odata.nextLink'
        } while ($uri)
        
        return $allFiles
    }
    catch {
        Write-Error "Failed to enumerate library files: $_"
        throw
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
        # Get sensitivity label using Graph API
        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$($File.id)?`$select=id,name,sensitivityLabel"
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET
        
        if ($response.sensitivityLabel) {
            return @{
                LabelId   = $response.sensitivityLabel.id
                LabelName = $response.sensitivityLabel.displayName
            }
        }
        
        # Fallback: Try extractSensitivityLabels endpoint
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

#region Main Processing
function Start-LabelMigration {
    param(
        [string]$SiteUrl,
        [string]$LibraryName,
        [string]$SourceLabelId,
        [string]$TargetLabelId,
        [string]$LogPath,
        [string]$ClientId,
        [string]$TenantId,
        [string]$ClientSecret
    )
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "SharePoint Sensitivity Label Migration" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Site URL: $SiteUrl" -ForegroundColor White
    Write-Host "Library: $LibraryName" -ForegroundColor White
    Write-Host "Source Label ID: $SourceLabelId" -ForegroundColor Yellow
    Write-Host "Target Label ID: $TargetLabelId" -ForegroundColor Green
    Write-Host ""
    
    # Initialize log file
    $logFile = Initialize-LogFile -Path $LogPath
    
    # Statistics
    $stats = @{
        TotalFiles      = 0
        FilesScanned    = 0
        FilesMatched    = 0
        FilesUpdated    = 0
        FilesFailed     = 0
        FilesSkipped    = 0
    }
    
    $result = $null
    try {
        # Get Graph access token
        Write-Host "Acquiring access token..." -ForegroundColor Yellow
        $accessToken = Get-GraphAccessToken -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret
        Write-Host "Access token acquired successfully" -ForegroundColor Green
        
        # Get site and drive IDs
        Write-Host "Resolving site and document library..." -ForegroundColor Yellow
        $siteInfo = Get-SiteAndDriveIds -SiteUrl $SiteUrl -LibraryName $LibraryName -AccessToken $accessToken
        Write-Host "Site ID: $($siteInfo.SiteId)" -ForegroundColor Gray
        Write-Host "Drive ID: $($siteInfo.DriveId)" -ForegroundColor Gray
        
        # Get all files in the library
        Write-Host "Enumerating files in document library..." -ForegroundColor Yellow
        $files = Get-AllLibraryFiles -SiteId $siteInfo.SiteId -DriveId $siteInfo.DriveId -AccessToken $accessToken
        $stats.TotalFiles = $files.Count
        Write-Host "Found $($files.Count) files to scan" -ForegroundColor Green
        Write-Host ""
        
        # Process each file
        $progress = 0
        foreach ($file in $files) {
            $progress++
            $percentComplete = [math]::Round(($progress / $stats.TotalFiles) * 100, 1)
            Write-Progress -Activity "Scanning files for sensitivity labels" -Status "$progress of $($stats.TotalFiles) - $($file.name)" -PercentComplete $percentComplete
            
            $stats.FilesScanned++
            
            # Get current label
            $labelInfo = Get-FileLabelInfo -File $file -SiteId $siteInfo.SiteId -DriveId $siteInfo.DriveId -AccessToken $accessToken
            
            if (-not $labelInfo) {
                Write-Verbose "No label found for: $($file.name)"
                continue
            }
            
            # Check if label matches source
            if ($labelInfo.LabelId -eq $SourceLabelId) {
                $stats.FilesMatched++
                Write-Host "Match found: $($file.name)" -ForegroundColor Cyan
                
                $logEntry = @{
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
                    # Update the sensitivity label
                    $updateResult = Set-FileSensitivityLabelViaGraph `
                        -SiteId $siteInfo.SiteId `
                        -DriveId $siteInfo.DriveId `
                        -ItemId $file.id `
                        -LabelId $TargetLabelId `
                        -AccessToken $accessToken
                    
                    if ($updateResult) {
                        $stats.FilesUpdated++
                        $logEntry.Status = "Updated"
                        $logEntry.NewLabelName = "Target Label Applied"
                        Write-Host "  Updated successfully" -ForegroundColor Green
                    }
                    else {
                        $stats.FilesFailed++
                        $logEntry.Status = "Failed"
                        $logEntry.ErrorMessage = "Update operation returned false"
                        Write-Host "  Update failed" -ForegroundColor Red
                    }
                }
                catch {
                    $stats.FilesFailed++
                    $logEntry.Status = "Error"
                    $logEntry.ErrorMessage = $_.Exception.Message
                    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
                }
                
                Write-LogEntry -LogPath $logFile -Entry $logEntry
            }
        }
        
        Write-Progress -Activity "Scanning files for sensitivity labels" -Completed
        $result = $stats
    }
    catch {
        Write-Error "Migration process failed: $_"
        throw
    }
    finally {
        # Display summary
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "Migration Summary" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "Total files in library: $($stats.TotalFiles)" -ForegroundColor White
        Write-Host "Files scanned: $($stats.FilesScanned)" -ForegroundColor White
        Write-Host "Files matching source label: $($stats.FilesMatched)" -ForegroundColor Yellow
        Write-Host "Files successfully updated: $($stats.FilesUpdated)" -ForegroundColor Green
        Write-Host "Files failed to update: $($stats.FilesFailed)" -ForegroundColor Red
        Write-Host ""
        Write-Host "Log file: $logFile" -ForegroundColor Cyan
        Write-Host ""
    }
    
    return $result
}
#endregion

#region Script Entry Point
try {
    # Verify required modules
    Install-RequiredModules
    
    # Initialize MIP context (optional - uses Graph API fallback if not available)
    $mipAvailable = Initialize-MipContext -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret
    
    # Start the migration process
    $results = Start-LabelMigration `
        -SiteUrl $SiteUrl `
        -LibraryName $LibraryName `
        -SourceLabelId $SourceLabelId `
        -TargetLabelId $TargetLabelId `
        -LogPath $LogPath `
        -ClientId $ClientId `
        -TenantId $TenantId `
        -ClientSecret $ClientSecret
    
    # Exit with appropriate code
    if ($results.FilesFailed -gt 0) {
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