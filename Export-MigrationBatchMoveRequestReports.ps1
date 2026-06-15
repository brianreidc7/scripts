<#
.SYNOPSIS
    Exports move request statistics and report for every mailbox in an Exchange Online migration batch.

.DESCRIPTION
    For each mailbox in a migration batch, retrieves the move request statistics (including the
    detailed migration report) and exports two files per mailbox:
      - A CSV containing all statistics properties
      - A CSV containing the full migration report log entries

    Completed migrations are named:    <mailbox>-<yyyyMMdd-HHmmss>-Statistics.csv / -Report.csv
    In-progress (or any non-completed) migrations are named:
                                        <mailbox>-inProgress-Statistics.csv / -Report.csv

.PARAMETER BatchName
    The name of the migration batch to process. If omitted, all migration batches are processed.

.PARAMETER OutputPath
    The folder path where export files will be saved. Defaults to the current directory.
    Sub-folders are created per batch when more than one batch is processed.

.EXAMPLE
    .\Export-MigrationBatchMoveRequestReports.ps1 -BatchName "Batch-Wave1" -OutputPath "C:\MigrationReports"

.EXAMPLE
    .\Export-MigrationBatchMoveRequestReports.ps1 -OutputPath "C:\MigrationReports"
    Processes all migration batches.

.NOTES
    Requires an active Exchange Online connection (Connect-ExchangeOnline).
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, HelpMessage = 'Name of the migration batch to process. Omit to process all batches.')]
    [string]$BatchName,

    [Parameter(Mandatory = $false, HelpMessage = 'Output folder for the exported files.')]
    [string]$OutputPath = (Get-Location).Path
)

#region Helpers

function Get-SafeFileName {
    param ([string]$Name)
    return $Name -replace '[<>:"/\\|?*\x00-\x1F]', '_'
}

#endregion

# ── Ensure output directory exists ──────────────────────────────────────────
if (-not (Test-Path -Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Write-Verbose "Created output directory: $OutputPath"
}

# ── Retrieve migration batches ───────────────────────────────────────────────
try {
    if ($PSBoundParameters.ContainsKey('BatchName')) {
        $batches = @(Get-MigrationBatch -Identity $BatchName -ErrorAction Stop)
    }
    else {
        $batches = @(Get-MigrationBatch -ErrorAction Stop)
    }
}
catch {
    Write-Error "Failed to retrieve migration batch(es): $_"
    return
}

if ($batches.Count -eq 0) {
    Write-Warning 'No migration batches found.'
    return
}

Write-Host "Found $($batches.Count) batch(es) to process." -ForegroundColor Cyan

$totalExported = 0
$totalFailed   = 0

foreach ($batch in $batches) {

    $batchName = $batch.Identity.ToString()
    Write-Host "`nBatch: $batchName" -ForegroundColor Cyan

    # When processing multiple batches, isolate each in its own sub-folder
    $batchOutputPath = if ($batches.Count -gt 1) {
        $subFolder = Join-Path -Path $OutputPath -ChildPath (Get-SafeFileName $batchName)
        if (-not (Test-Path -Path $subFolder)) {
            New-Item -ItemType Directory -Path $subFolder -Force | Out-Null
        }
        $subFolder
    }
    else {
        $OutputPath
    }

    # ── Get all users in this batch via Get-MigrationUser ─────────────────────
    # (Get-MoveRequest -BatchName is unreliable in Exchange Online; Get-MigrationUser
    #  -BatchId is the authoritative way to enumerate batch members.)
    try {
        $migrationUsers = @(Get-MigrationUser -BatchId $batchName -ResultSize Unlimited -ErrorAction Stop)
    }
    catch {
        Write-Warning "Could not retrieve migration users for batch '$batchName': $_"
        continue
    }

    if ($migrationUsers.Count -eq 0) {
        Write-Warning "No migration users found in batch '$batchName'."
        continue
    }

    Write-Host "  $($migrationUsers.Count) migration user(s) found." -ForegroundColor Gray

    foreach ($migUser in $migrationUsers) {

        # Use the user's primary address / identity as the mailbox identifier
        $mailboxId = $migUser.Identity.ToString()

        Write-Host "  Processing: $mailboxId" -ForegroundColor Yellow

        try {
            $stats = Get-MoveRequestStatistics -Identity $mailboxId -IncludeReport -ErrorAction Stop

            # ── Build the file-name base ────────────────────────────────────
            $completedStatuses = @('Completed', 'CompletedWithWarning')

            if ($stats.Status -in $completedStatuses -and $stats.CompletionTimestamp) {
                $dateSuffix = ([datetime]$stats.CompletionTimestamp).ToString('yyyyMMdd-HHmmss')
                $fileBase   = Get-SafeFileName "$mailboxId-$dateSuffix"
            }
            else {
                $fileBase = Get-SafeFileName "$mailboxId-inProgress"
            }

            # ── Export statistics (all scalar properties, no Report blob) ───
            $statsPath = Join-Path -Path $batchOutputPath -ChildPath "$fileBase-Statistics.csv"
            $stats |
                Select-Object -ExcludeProperty Report |
                Export-Csv -Path $statsPath -NoTypeInformation -Encoding UTF8 -Force
            Write-Verbose "    Statistics -> $statsPath"

            # ── Export report log entries ────────────────────────────────────
            if ($stats.Report -and $stats.Report.Entries) {
                $reportPath = Join-Path -Path $batchOutputPath -ChildPath "$fileBase-Report.csv"
                $stats.Report.Entries |
                    Select-Object -Property CreationTime, Type, Description |
                    Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8 -Force
                Write-Verbose "    Report     -> $reportPath"
            }

            Write-Host "    Exported: $fileBase" -ForegroundColor Green
            $totalExported++
        }
        catch {
            Write-Warning "Failed to export data for '$mailboxId': $_"
            $totalFailed++
        }
    }
}

# ── Summary ──────────────────────────────────────────────────────────────────
$summaryColour = if ($totalFailed -gt 0) { 'Yellow' } else { 'Green' }
Write-Host "`nComplete. Exported: $totalExported | Failed: $totalFailed" -ForegroundColor $summaryColour
Write-Host "Output folder: $OutputPath" -ForegroundColor Green
