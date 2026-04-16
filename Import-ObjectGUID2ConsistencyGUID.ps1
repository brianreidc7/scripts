<#
.SYNOPSIS
    Imports ObjectGUID mappings into a destination directory, setting msDSConsistencyGUID.

.DESCRIPTION
    This script imports ObjectGUID mappings from a CSV file and matches objects in 
    the destination directory by PrimarySMTPAddress, then sets their msDSConsistencyGUID 
    attribute to the source ObjectGUID

.PARAMETER DestinationDirectory
    The destination directory path or identifier where objects will be matched and updated

.PARAMETER ExportPath
    The file path to the CSV mapping file to import (default: .\ObjectGUID_Mapping.csv)

.EXAMPLE
    .\Import-ObjectGUID2ConsistencyGUID.ps1 -DestinationDirectory "DC=contoso,DC=com"

.EXAMPLE
    .\Import-ObjectGUID2ConsistencyGUID.ps1 -DestinationDirectory "DC=contoso,DC=com" -ExportPath "C:\temp\mapping.csv"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$DestinationDirectory,
    
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = ".\ObjectGUID_Mapping.csv"
)

# ===== IMPORT REQUIRED MODULE =====
try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    Write-Error "Failed to import ActiveDirectory module. Please ensure the AD PowerShell module is installed."
    exit 1
}

# ===== CONFIGURATION =====
$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

Write-Host "================================================================================"
Write-Host "ObjectGUID & PrimarySMTPAddress Import Tool"
Write-Host "================================================================================"
Write-Host ""

# ===== IMPORT PHASE =====
if (-not (Test-Path $ExportPath)) {
    Write-Error "Import file not found at: $ExportPath"
    exit 1
}

Write-Host "PHASE 1: IMPORT & MATCH" -ForegroundColor Cyan
Write-Host "Destination Directory: $DestinationDirectory"
Write-Host "Source CSV: $ExportPath"
Write-Host ""

try {
    # Import the mapping
    Write-Verbose "Importing ObjectGUID mapping from CSV..."
    $mappingData = Import-Csv -Path $ExportPath -Encoding UTF8
    
    if ($null -eq $mappingData) {
        Write-Error "No data found in CSV file"
        exit 1
    }
    
    # Normalize to array (single object imports as object, not array)
    if ($mappingData -is [object] -and $mappingData -isnot [array]) {
        $mappingData = @($mappingData)
    }

    Write-Host "Loaded $($mappingData.Count) mapping record(s)" -ForegroundColor Green
    Write-Host ""

    # Match and process objects in destination
    Write-Verbose "Retrieving AD User objects from destination: $DestinationDirectory..."
    $destObjects = Get-ADUser -Filter * -SearchBase $DestinationDirectory -Properties mail, sAMAccountName, "ms-DS-ConsistencyGUID" -ErrorAction SilentlyContinue
    
    Write-Host "Found $($destObjects.Count) object(s) in destination directory" -ForegroundColor Green
    Write-Host ""

    # Matching logic
    $matchedCount = 0
    $unmatchedCount = 0
    $updatedCount = 0
    $updateErrorCount = 0
    $results = @()

    foreach ($mapping in $mappingData) {
        $primarySmtp = $mapping.PrimarySMTPAddress
        $originalGuid = $mapping.ObjectGUID
        
        Write-Verbose "Processing: $primarySmtp"
        
        # Find matching object in destination by PrimarySMTPAddress
        $destObject = $destObjects | Where-Object { $_.mail -eq $primarySmtp }
        
        if ($destObject) {
            $matchedCount++
            $msDSConsistencyGUIDStatus = "PENDING"
            $msDSConsistencyGUIDError = $null

            # Attempt to set msDSConsistencyGUID attribute
            try {
                Write-Verbose "Attempting to set msDSConsistencyGUID for $primarySmtp"
                
                # Convert the source GUID to binary format for msDSConsistencyGUID
                if ($originalGuid -is [guid]) {
                    $guidBytes = $originalGuid.ToByteArray()
                }
                else {
                    $guidBytes = [guid]::Parse($originalGuid).ToByteArray()
                }

                # Check if ms-DS-ConsistencyGUID already has a value
                $currentValue = $destObject."ms-DS-ConsistencyGUID"
                
                if ($null -ne $currentValue -and $currentValue -ne $guidBytes) {
                    # Convert current value to compare
                    $currentGuid = [guid]$currentValue
                    $incomingGuid = [guid]$guidBytes
                    
                    if ($currentGuid -ne $incomingGuid) {
                        $msDSConsistencyGUIDStatus = "FAILED"
                        $msDSConsistencyGUIDError = "Attribute already contains a different value: $currentGuid. Will not overwrite with $incomingGuid"
                        $updateErrorCount++
                        Write-Host "CONFLICT: $primarySmtp - Attribute already has value $currentGuid, not overwriting with $incomingGuid" -ForegroundColor Yellow
                    }
                    else {
                        # Values are the same, mark as success
                        $msDSConsistencyGUIDStatus = "SUCCESS"
                        $updatedCount++
                        Write-Host "MATCHED: $primarySmtp -> msDSConsistencyGUID already set correctly" -ForegroundColor Green
                    }
                }
                else {
                    # Attribute is empty or doesn't exist, proceed with update
                    Set-ADObject -Identity $destObject.DistinguishedName -Replace @{ "ms-DS-ConsistencyGUID" = $guidBytes } -ErrorAction Stop
                    
                    $msDSConsistencyGUIDStatus = "SUCCESS"
                    $updatedCount++
                    Write-Host "MATCHED AND UPDATED: $primarySmtp -> msDSConsistencyGUID set" -ForegroundColor Green
                }
            }
            catch {
                $msDSConsistencyGUIDStatus = "FAILED"
                $msDSConsistencyGUIDError = $_.Exception.Message
                $updateErrorCount++
                Write-Host "MATCHED BUT UPDATE FAILED: $primarySmtp - $_" -ForegroundColor Yellow
            }

            $result = [PSCustomObject]@{
                PrimarySMTPAddress        = $primarySmtp
                SourceObjectGUID          = $originalGuid
                DestinationObjectGUID     = $destObject.ObjectGUID
                DestinationDisplayName    = $destObject.DisplayName
                msDSConsistencyGUIDStatus = $msDSConsistencyGUIDStatus
                msDSConsistencyGUIDError  = $msDSConsistencyGUIDError
                Status                    = "MATCHED"
                Timestamp                 = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        else {
            $unmatchedCount++
            $result = [PSCustomObject]@{
                PrimarySMTPAddress        = $primarySmtp
                SourceObjectGUID          = $originalGuid
                DestinationObjectGUID     = $null
                DestinationDisplayName    = $null
                msDSConsistencyGUIDStatus = "N/A"
                msDSConsistencyGUIDError  = $null
                Status                    = "NOT_FOUND"
                Timestamp                 = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            Write-Host "NOT FOUND: $primarySmtp" -ForegroundColor Yellow
        }
        
        $results += $result
    }

    Write-Host ""
    Write-Host "=" * 80
    Write-Host "IMPORT RESULTS" -ForegroundColor Cyan
    Write-Host "=" * 80
    Write-Host "Matched:               $matchedCount"
    Write-Host "Successfully Updated:  $updatedCount"
    Write-Host "Update Errors:         $updateErrorCount"
    Write-Host "Not Found:             $unmatchedCount"
    Write-Host "Total:                 $($results.Count)"
    Write-Host ""

    # Export matching results
    $resultsPath = ($ExportPath -replace '\.csv$', '') + "_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $results | Export-Csv -Path $resultsPath -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "Results exported to: $resultsPath" -ForegroundColor Green
    
    # Summary table
    Write-Host ""
    Write-Host "MATCHED OBJECTS:" -ForegroundColor Green
    $results | Where-Object { $_.Status -eq "MATCHED" } | Format-Table -Property PrimarySMTPAddress, SourceObjectGUID, DestinationObjectGUID, msDSConsistencyGUIDStatus -AutoSize

    if ($updateErrorCount -gt 0) {
        Write-Host ""
        Write-Host "UPDATE ERRORS:" -ForegroundColor Yellow
        $results | Where-Object { $_.msDSConsistencyGUIDStatus -eq "FAILED" } | Format-Table -Property PrimarySMTPAddress, msDSConsistencyGUIDError -AutoSize
    }

    if ($unmatchedCount -gt 0) {
        Write-Host ""
        Write-Host "UNMATCHED OBJECTS:" -ForegroundColor Yellow
        $results | Where-Object { $_.Status -eq "NOT_FOUND" } | Format-Table -Property PrimarySMTPAddress, SourceObjectGUID -AutoSize
    }
}
catch {
    Write-Error "Error during import phase: $_"
    exit 1
}

Write-Host ""
Write-Host "Script completed successfully!" -ForegroundColor Green
