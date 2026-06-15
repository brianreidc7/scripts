<#
.SYNOPSIS
    Exports move request statistics and report for every mailbox in an Exchange Online migration batch.

.DESCRIPTION
    For each mailbox in a migration batch, retrieves the move request statistics (including the
    detailed migration report) and exports two files per mailbox:
      - A CSV containing all statistics properties
      - A TXT file containing the full migration report log entries

    Completed migrations are named:    <mailbox>-<yyyyMMdd-HHmmss>-Statistics.csv / -Report.txt
    In-progress (or any non-completed) migrations are named:
                                        <mailbox>-inProgress-Statistics.csv / -Report.txt

.PARAMETER BatchName
    The name of the migration batch to process. If omitted, all migration batches are processed.

.PARAMETER OutputPath
    The folder path where export files will be saved. Defaults to C:\temp.
    Sub-folders are created per batch when more than one batch is processed.

.PARAMETER RemoveInProgressFiles
    When a migration is now complete, automatically delete any existing -inProgress files
    for that mailbox before writing the new dated files.

.EXAMPLE
    .\Export-MigrationBatchMoveRequestReports.ps1 -BatchName "Batch-Wave1"
    Exports to the default folder C:\temp.

.EXAMPLE
    .\Export-MigrationBatchMoveRequestReports.ps1 -BatchName "Batch-Wave1" -RemoveInProgressFiles
    Exports to C:\temp and removes any stale -inProgress files for completed mailboxes.

.EXAMPLE
    .\Export-MigrationBatchMoveRequestReports.ps1 -BatchName "Batch-Wave1" -OutputPath "C:\MigrationReports"
    Exports to a custom folder.

.EXAMPLE
    .\Export-MigrationBatchMoveRequestReports.ps1 -OutputPath "C:\MigrationReports"
    Processes all migration batches into a custom folder.

.NOTES
    Requires an active Exchange Online connection (Connect-ExchangeOnline).
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, HelpMessage = 'Name of the migration batch to process. Omit to process all batches.')]
    [string]$BatchName,

    [Parameter(Mandatory = $false, HelpMessage = 'Output folder for the exported files. Defaults to C:\temp.')]
    [string]$OutputPath = 'C:\temp',

    [Parameter(Mandatory = $false, HelpMessage = 'Delete stale -inProgress files when a migration is found to be complete.')]
    [switch]$RemoveInProgressFiles
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

    # Always create a sub-folder named after the batch
    $batchOutputPath = Join-Path -Path $OutputPath -ChildPath (Get-SafeFileName $batchName)
    if (-not (Test-Path -Path $batchOutputPath)) {
        New-Item -ItemType Directory -Path $batchOutputPath -Force | Out-Null
        Write-Verbose "Created batch folder: $batchOutputPath"
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

    # Collect stats for all mailboxes so we can write a batch-level summary
    $batchStatsList = [System.Collections.Generic.List[object]]::new()

    foreach ($migUser in $migrationUsers) {

        # Use the user's primary address / identity as the mailbox identifier
        $mailboxId = $migUser.Identity.ToString()

        Write-Host "  Processing: $mailboxId" -ForegroundColor Yellow

        try {
            $stats = Get-MoveRequestStatistics -Identity $mailboxId -IncludeReport -ErrorAction Stop
            Write-Verbose "       Status: $($stats.Status.ToString()) | CompletionTimestamp: $($stats.CompletionTimestamp)"

            # ── Build the file-name base ────────────────────────────────────
            $completedStatuses = @('Completed', 'CompletedWithWarning')
            $statusString = $stats.Status.ToString()

            $safeMailboxId = Get-SafeFileName $mailboxId

            if ($statusString -in $completedStatuses -and $stats.CompletionTimestamp) {
                $dateSuffix = ([datetime]$stats.CompletionTimestamp).ToString('yyyyMMdd-HHmmss')
                $fileBase   = "$safeMailboxId-$dateSuffix"

                # ── Remove stale -inProgress files if switch is set ─────────
                # Matches: -Statistics.csv, -Report.txt, -BadItemsHistory.txt
                if ($RemoveInProgressFiles) {
                    $staleFiles = Get-ChildItem -Path $batchOutputPath -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -like "$safeMailboxId-inProgress-*" }
                    foreach ($stale in $staleFiles) {
                        Remove-Item -Path $stale.FullName -Force
                        Write-Host "    Removed stale file: $($stale.Name)" -ForegroundColor DarkYellow
                    }
                }
            }
            else {
                $fileBase = "$safeMailboxId-inProgress"
            }

            # ── Export statistics (all scalar properties, no Report blob) ───
            $statsPath = Join-Path -Path $batchOutputPath -ChildPath "$fileBase-Statistics.csv"

            if (Test-Path -Path $statsPath) {
                Write-Host "    Skipped (already exists): $fileBase" -ForegroundColor DarkGray
                continue
            }

            $stats |
                Select-Object -ExcludeProperty Report,RolloutNames |
                Export-Csv -Path $statsPath -NoTypeInformation -Encoding UTF8 -Force
            Write-Verbose "       Statistics -> $statsPath"

            # Accumulate for batch summary (exclude Report blob)
            $batchStatsList.Add(($stats | Select-Object -ExcludeProperty Report,RolloutNames))

            # ── Export report log entries ────────────────────────────────────
            if ($stats.Report -and $stats.Report.Entries) {
                $reportPath = Join-Path -Path $batchOutputPath -ChildPath "$fileBase-Report.txt"
                $stats |
                    Select-Object -Property Report | Format-List |
                    Out-File -Path $reportPath -Encoding UTF8 -Force
                Write-Verbose "       Report     -> $reportPath"
            }

            # ── Export bad items history if BadItems > 0 (any status) ─────
            if ($stats.BadItemsEncountered -gt 0 -and
                $stats.Report -and $stats.Report.BadItemsHistory) {

                $badItemsPath = Join-Path -Path $batchOutputPath -ChildPath "$fileBase-BadItemsHistory.txt"
                $stats.Report.BadItemsHistory |
                    Out-File -Path $badItemsPath -Encoding UTF8 -Force
                Write-Verbose "       BadItemsHistory -> $badItemsPath"
                Write-Host "    Bad items history exported ($($stats.BadItemsEncountered) bad item(s)): $(Split-Path $badItemsPath -Leaf)" -ForegroundColor Yellow
            }

            Write-Host "    Exported: $fileBase" -ForegroundColor Green
            $totalExported++
        }
        catch {
            Write-Warning "Failed to export data for '$mailboxId': $_"
            $totalFailed++
        }
    }

    # ── Batch summary file (always replaced) ────────────────────────────────
    if ($batchStatsList.Count -gt 0) {
        $safeBatchName    = Get-SafeFileName $batchName
        $batchSummaryPath = Join-Path -Path $batchOutputPath -ChildPath "$safeBatchName-BatchSummary.csv"

        if (Test-Path -Path $batchSummaryPath) {
            Remove-Item -Path $batchSummaryPath -Force
            Write-Verbose "    Deleted existing batch summary: $batchSummaryPath"
        }

        $batchStatsList |
            Export-Csv -Path $batchSummaryPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "  Batch summary written: $(Split-Path $batchSummaryPath -Leaf)" -ForegroundColor Cyan
    }
}

# ── Summary ──────────────────────────────────────────────────────────────────
$summaryColour = if ($totalFailed -gt 0) { 'Yellow' } else { 'Green' }
Write-Host "`nComplete. Exported: $totalExported | Failed: $totalFailed" -ForegroundColor $summaryColour
Write-Host "Output folder: $OutputPath" -ForegroundColor Green
