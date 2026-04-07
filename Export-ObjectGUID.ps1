<#
.SYNOPSIS
    Exports ObjectGUID and PrimarySMTPAddress from a source directory.

.DESCRIPTION
    This script exports ObjectGUID and mail (SMTP address) from Active Directory objects
    into a CSV file for later use.

.PARAMETER SourceDirectory
    The source directory path (e.g., "OU=Users,DC=contoso,DC=com") or a searchbase for AD queries

.PARAMETER ExportPath
    The file path where the CSV mapping will be saved (default: .\ObjectGUID_Mapping.csv)

.EXAMPLE
    .\Export-ObjectGUID.ps1 -SourceDirectory "OU=Users,DC=contoso,DC=com"

.EXAMPLE
    .\Export-ObjectGUID.ps1 -SourceDirectory "OU=Users,DC=contoso,DC=com" -ExportPath "C:\temp\mapping.csv"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDirectory,
    
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

Write-Host "=" * 80
Write-Host "ObjectGUID & PrimarySMTPAddress Export Tool"
Write-Host "=" * 80
Write-Host ""

# ===== EXPORT PHASE =====
Write-Host "PHASE 1: EXPORT" -ForegroundColor Cyan
Write-Host "Source Directory: $SourceDirectory"
Write-Host "Export Path: $ExportPath"
Write-Host ""

try {
    Write-Verbose "Retrieving AD User objects from source: $SourceDirectory"
    
    # Retrieve User objects from AD
    Write-Verbose "Querying for User objects..."
    $sourceObjects = Get-ADUser -Filter * -SearchBase $SourceDirectory -Properties mail, sAMAccountName -ErrorAction SilentlyContinue | 
        Where-Object { $_.mail -ne $null } |
        Select-Object -Property @{Name = "ObjectGUID"; Expression = { $_.ObjectGUID } },
                                @{Name = "PrimarySMTPAddress"; Expression = { $_.mail } },
                                DisplayName
    
    if ($sourceObjects.Count -eq 0) {
        Write-Warning "No objects with mail addresses found in: $SourceDirectory"
    }
    else {
        Write-Host "Found $($sourceObjects.Count) object(s) with mail addresses" -ForegroundColor Green
    }

    # Export to CSV
    $sourceObjects | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "Successfully exported mapping to: $ExportPath" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Error "Error during export phase: $_"
    exit 1
}

Write-Host ""
Write-Host "Export completed successfully!" -ForegroundColor Green
