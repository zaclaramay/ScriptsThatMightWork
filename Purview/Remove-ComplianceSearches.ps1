# Remove Compliance Searches from CSV File
# This script connects to Security & Compliance Center and removes compliance searches listed in a CSV file

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvFilePath,
    
    [Parameter(Mandatory=$false)]
    [switch]$WhatIf = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$Force = $false
)

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Function to log actions
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage
    
    # Optional: Write to log file
    # $logMessage | Out-File -FilePath "compliance_search_removal.log" -Append
}

try {
    # Check if CSV file exists
    if (-not (Test-Path $CsvFilePath)) {
        Write-ColorOutput "ERROR: CSV file not found at path: $CsvFilePath" "Red"
        exit 1
    }

    # Import CSV file
    Write-Log "Reading CSV file: $CsvFilePath"
    $searchList = Import-Csv $CsvFilePath

    # Validate CSV structure (expecting a column named 'SearchName' or 'Name')
    $searchNameColumn = $null
    if ($searchList[0].PSObject.Properties.Name -contains 'SearchName') {
        $searchNameColumn = 'SearchName'
    } elseif ($searchList[0].PSObject.Properties.Name -contains 'Name') {
        $searchNameColumn = 'Name'
    } else {
        Write-ColorOutput "ERROR: CSV file must contain either 'SearchName' or 'Name' column" "Red"
        Write-ColorOutput "Available columns: $($searchList[0].PSObject.Properties.Name -join ', ')" "Yellow"
        exit 1
    }

    Write-Log "Found $($searchList.Count) compliance searches to process"
    Write-Log "Using column: $searchNameColumn"

    # Check if Security & Compliance Center PowerShell module is available
    Write-Log "Checking Security & Compliance Center connection..."
    
    try {
        # Test connection by trying to get compliance searches
        $testSearch = Get-ComplianceSearch -ResultSize 1 -ErrorAction Stop
        Write-ColorOutput "Connected to Security & Compliance Center" "Green"
    }
    catch {
        Write-ColorOutput "ERROR: Not connected to Security & Compliance Center PowerShell" "Red"
        Write-ColorOutput "Please run: Connect-IPPSSession" "Yellow"
        exit 1
    }

    # Process each search
    $successCount = 0
    $failureCount = 0
    $skippedCount = 0

    foreach ($item in $searchList) {
        $searchName = $item.$searchNameColumn
        
        if ([string]::IsNullOrWhiteSpace($searchName)) {
            Write-ColorOutput "WARNING: Skipping empty search name" "Yellow"
            $skippedCount++
            continue
        }

        Write-Log "Processing: $searchName"

        try {
            # Check if compliance search exists
            $existingSearch = Get-ComplianceSearch -Identity $searchName -ErrorAction SilentlyContinue
            
            if (-not $existingSearch) {
                Write-ColorOutput "WARNING: Compliance search '$searchName' not found - skipping" "Yellow"
                $skippedCount++
                continue
            }

            # Check if search is running
            if ($existingSearch.Status -eq "InProgress") {
                Write-ColorOutput "WARNING: Search '$searchName' is currently running. Cannot delete." "Yellow"
                $skippedCount++
                continue
            }

            if ($WhatIf) {
                Write-ColorOutput "WHAT IF: Would remove compliance search '$searchName'" "Cyan"
                $successCount++
            } else {
                # Remove the compliance search
                if ($Force) {
                    Remove-ComplianceSearch -Identity $searchName -Confirm:$false -ErrorAction Stop
                } else {
                    Remove-ComplianceSearch -Identity $searchName -ErrorAction Stop
                }
                
                Write-ColorOutput "SUCCESS: Removed compliance search '$searchName'" "Green"
                $successCount++
            }
        }
        catch {
            Write-ColorOutput "ERROR: Failed to remove '$searchName' - $($_.Exception.Message)" "Red"
            $failureCount++
        }
    }

    # Summary
    Write-Log "=== SUMMARY ==="
    Write-ColorOutput "Total processed: $($searchList.Count)" "White"
    Write-ColorOutput "Successful: $successCount" "Green"
    Write-ColorOutput "Failed: $failureCount" "Red"
    Write-ColorOutput "Skipped: $skippedCount" "Yellow"

    if ($WhatIf) {
        Write-ColorOutput "Run without -WhatIf to actually remove the searches" "Cyan"
    }
}
catch {
    Write-ColorOutput "FATAL ERROR: $($_.Exception.Message)" "Red"
    exit 1
}

# Example CSV format:
<#
SearchName
"HR Investigation 2024"
"Legal Hold Search 1"
"Compliance Audit Q1"
#>

# Usage examples:
# .\Remove-ComplianceSearches.ps1 -CsvFilePath "C:\searches.csv" -WhatIf
# .\Remove-ComplianceSearches.ps1 -CsvFilePath "C:\searches.csv" -Force
# .\Remove-ComplianceSearches.ps1 -CsvFilePath "C:\searches.csv"