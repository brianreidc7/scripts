#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Compares the Standard Preset Security Policy in Microsoft 365 to a duplicated custom policy.

.DESCRIPTION
    For each policy type bundled within the Standard Preset Security Policy, this script
    retrieves both the live preset policy and a previously duplicated custom policy and
    reports any differences in their settings:
      - Anti-spam (HostedContentFilterPolicy)
      - Anti-malware (MalwareFilterPolicy)
      - Anti-phishing (AntiPhishPolicy)
      - Safe Links (SafeLinksPolicy)         [requires Defender for Office 365 Plan 1 or 2]
      - Safe Attachments (SafeAttachmentsPolicy) [requires Defender for Office 365 Plan 1 or 2]

    Only the settings that would normally be copied when duplicating are compared; metadata
    and read-only properties (identity, timestamps, GUIDs, etc.) are ignored. Use this to
    confirm a duplicate still matches the current preset, or to spot drift after the preset
    has been updated by Microsoft.

    This script is read-only and makes no changes.

.PARAMETER PolicyName
    The name of the duplicated policies to compare against the preset. Defaults to
    "Standard Preset Policy - Duplicate". The same name is used for all five policy types.

.PARAMETER PolicyTypes
    One or more policy types to compare. Valid values:
      AntiSpam, AntiMalware, AntiPhish, SafeLinks, SafeAttachments, All
    If omitted, an interactive menu is displayed.

.PARAMETER ShowIdentical
    Also list settings that match, not just the differences.

.EXAMPLE
    .\Compare-StandardPresetSecurityPolicy.ps1

.EXAMPLE
    .\Compare-StandardPresetSecurityPolicy.ps1 -PolicyName "Contoso Standard Policy"

.EXAMPLE
    .\Compare-StandardPresetSecurityPolicy.ps1 -PolicyTypes AntiSpam, AntiPhish

.EXAMPLE
    .\Compare-StandardPresetSecurityPolicy.ps1 -PolicyTypes All -ShowIdentical

.NOTES
    Prerequisites:
    - ExchangeOnlineManagement module v3.0 or later must be installed:
        Install-Module -Name ExchangeOnlineManagement
    - You must have appropriate admin permissions in Microsoft 365
      (Security Reader, Security Administrator or Global Administrator)
    - Script connects to Exchange Online if not already connected
#>

#Requires -Module ExchangeOnlineManagement

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$PolicyName = "Standard Preset Policy - Duplicate",

    # Which policy types to compare. If omitted, an interactive menu is shown.
    [Parameter(Mandatory = $false)]
    [ValidateSet('AntiSpam', 'AntiMalware', 'AntiPhish', 'SafeLinks', 'SafeAttachments', 'All')]
    [string[]]$PolicyTypes,

    # Also list settings that are identical, not just the differences.
    [Parameter(Mandatory = $false)]
    [switch]$ShowIdentical
)

#region --- Connection ---

try {
    $null = Get-OrganizationConfig -ErrorAction Stop
    Write-Host "Connected to Exchange Online." -ForegroundColor Green
}
catch {
    Write-Host "Not connected to Exchange Online. Initiating connection..." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowProgress $true
}

#endregion

#region --- Helper Functions ---

# Builds a hashtable of comparable settings from a policy object, stripping metadata and
# read-only properties. Mirrors the property set used when duplicating a preset policy.
function Get-ComparableParams {
    param (
        [Parameter(Mandatory)]
        [object]$Policy,

        [string[]]$ExcludeProperties = @(
            'ZapEnabled', 'EnableSuspiciousSafetyTip', 'PolicyTag', 'EnableOrganizationBranding',
            'EnableBlockingEncryptedAttachments', 'QuarantineTagForBlockingEncryptedAttachments', 'DirectoryObjectVersion'
        )
    )

    $metadataProps = @(
        'Identity', 'Id', 'ExchangeObjectId', 'Guid', 'DistinguishedName', 'Name',
        'WhenCreated', 'WhenChanged', 'WhenCreatedUTC', 'WhenChangedUTC',
        'ObjectState', 'IsDefault', 'IsBuiltIn', 'IsPreset', 'RecommendedPolicyType',
        'ExchangeVersion', 'OrganizationId', 'PSShowComputerName', 'PSComputerName',
        'RunspaceId', 'AdminDisplayName', 'Rules', 'ObjectCategory', 'IsValid', 'ObjectClass',
        'OriginatingServer', 'OrganizationalUnitRoot', 'IsPolicyOverrideApplied', 'IsBuiltInProtection'
    )

    $skip = $metadataProps + $ExcludeProperties

    $params = @{}
    foreach ($prop in $Policy.PSObject.Properties) {
        if ($prop.Name -in $skip) { continue }
        $params[$prop.Name] = $prop.Value
    }

    return $params
}

# Renders a value for readable comparison output.
function Format-Value {
    param ([object]$Value)

    if ($null -eq $Value) { return '<null>' }
    if ($Value -is [System.Collections.ICollection] -and $Value -isnot [string]) {
        if ($Value.Count -eq 0) { return '<empty>' }
        return ($Value | ForEach-Object { "$_" }) -join '; '
    }
    return "$Value"
}

# Compares the preset source policy to the duplicate and prints the differences.
function Compare-PolicyPair {
    param (
        [Parameter(Mandatory)] [string]$Label,
        [Parameter(Mandatory)] [object]$Source,
        [Parameter(Mandatory)] [object]$Duplicate
    )

    $sourceParams = Get-ComparableParams -Policy $Source
    $dupParams = Get-ComparableParams -Policy $Duplicate

    $allKeys = ($sourceParams.Keys + $dupParams.Keys) | Sort-Object -Unique

    $differences = 0
    $identical = 0

    foreach ($key in $allKeys) {
        $sVal = Format-Value $sourceParams[$key]
        $dVal = Format-Value $dupParams[$key]

        if ($sVal -eq $dVal) {
            $identical++
            if ($ShowIdentical) {
                Write-Host ("    = {0}" -f $key) -ForegroundColor DarkGray
                Write-Host ("        {0}" -f $sVal) -ForegroundColor DarkGray
            }
        }
        else {
            $differences++
            Write-Host ("    ~ {0}" -f $key) -ForegroundColor Yellow
            Write-Host ("        Preset    : {0}" -f $sVal) -ForegroundColor Gray
            Write-Host ("        Duplicate : {0}" -f $dVal) -ForegroundColor White
        }
    }

    if ($differences -eq 0) {
        Write-Host "  No setting differences — duplicate matches the preset ($identical settings compared)." -ForegroundColor Green
    }
    else {
        Write-Host ("  {0} difference(s), {1} identical setting(s)." -f $differences, $identical) -ForegroundColor Yellow
    }

    return $differences
}

#endregion

#region --- Policy Type Selection Menu ---

$validTypes = @('AntiSpam', 'AntiMalware', 'AntiPhish', 'SafeLinks', 'SafeAttachments')

if (-not $PolicyTypes) {
    Write-Host ""
    Write-Host "Which policy types do you want to compare?" -ForegroundColor Cyan
    Write-Host "  [1] Anti-Spam"
    Write-Host "  [2] Anti-Malware"
    Write-Host "  [3] Anti-Phishing"
    Write-Host "  [4] Safe Links        (requires Defender for Office 365 Plan 1 or 2)"
    Write-Host "  [5] Safe Attachments  (requires Defender for Office 365 Plan 1 or 2)"
    Write-Host "  [A] All of the above"
    Write-Host ""

    $selection = Read-Host "Enter one or more numbers/letters separated by commas (e.g. 1,3 or A)"
    $tokens = $selection -split '[,\s]+' | ForEach-Object { $_.Trim().ToUpper() } | Where-Object { $_ -ne '' }

    if ('A' -in $tokens) {
        $PolicyTypes = $validTypes
    }
    else {
        $map = @{ '1' = 'AntiSpam'; '2' = 'AntiMalware'; '3' = 'AntiPhish'; '4' = 'SafeLinks'; '5' = 'SafeAttachments' }
        $PolicyTypes = foreach ($t in $tokens) {
            if ($map.ContainsKey($t)) { $map[$t] }
            else { Write-Warning "Unrecognised selection '$t' — ignoring." }
        }
    }

    if (-not $PolicyTypes) {
        Write-Error "No valid policy types selected. Exiting."
        return
    }
}
elseif ($PolicyTypes -contains 'All') {
    $PolicyTypes = $validTypes
}

#endregion

$results = [ordered]@{}

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  Comparing Standard Preset Security Policy to Duplicate" -ForegroundColor Cyan
Write-Host "  Duplicate name : '$PolicyName'" -ForegroundColor Cyan
Write-Host "  Policy types   : $($PolicyTypes -join ', ')" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan

# Maps each policy type to its Get cmdlet and the license-sensitive flag.
$policyMap = [ordered]@{
    AntiSpam        = @{ Label = 'Anti-Spam';        Cmdlet = 'Get-HostedContentFilterPolicy'; Requires = $false }
    AntiMalware     = @{ Label = 'Anti-Malware';     Cmdlet = 'Get-MalwareFilterPolicy';        Requires = $false }
    AntiPhish       = @{ Label = 'Anti-Phishing';    Cmdlet = 'Get-AntiPhishPolicy';            Requires = $false }
    SafeLinks       = @{ Label = 'Safe Links';       Cmdlet = 'Get-SafeLinksPolicy';            Requires = $true }
    SafeAttachments = @{ Label = 'Safe Attachments'; Cmdlet = 'Get-SafeAttachmentPolicy';       Requires = $true }
}

foreach ($type in $validTypes) {

    if ($type -notin $PolicyTypes) { continue }

    $info = $policyMap[$type]

    Write-Host ""
    Write-Host ("[{0}] {1}" -f $info.Label, $info.Cmdlet.Replace('Get-', '')) -ForegroundColor Yellow

    try {
        $sourcePresetName = (& $info.Cmdlet |
            Where-Object { $_.Name -like 'Standard Preset Security Policy*' } |
            Select-Object -First 1).Name
        if (-not $sourcePresetName) { throw "Standard Preset Security Policy not found for $($info.Label)." }
        $source = & $info.Cmdlet -Identity $sourcePresetName -ErrorAction Stop

        $duplicate = & $info.Cmdlet -Identity $PolicyName -ErrorAction SilentlyContinue
        if (-not $duplicate) {
            Write-Warning "  Duplicate policy '$PolicyName' not found — nothing to compare."
            $results[$type] = 'Missing (duplicate not found)'
            continue
        }

        $diffCount = Compare-PolicyPair -Label $info.Label -Source $source -Duplicate $duplicate
        $results[$type] = if ($diffCount -eq 0) { 'Match' } else { "$diffCount difference(s)" }
    }
    catch {
        $msg = $_.Exception.Message
        if ($info.Requires -and $msg -match 'recognized|licensed|subscription|not found') {
            Write-Warning "  $($info.Label) not available — requires Defender for Office 365 Plan 1 or 2."
            $results[$type] = 'Skipped (license not available)'
        }
        else {
            Write-Warning "  Failed to compare $($info.Label) policy: $msg"
            $results[$type] = "Failed: $msg"
        }
    }
}

#region --- Summary ---

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  Summary" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
foreach ($item in $results.GetEnumerator()) {
    $color = if ($item.Value -eq 'Match') { 'Green' }
             elseif ($item.Value -like 'Skipped*' -or $item.Value -like 'Missing*') { 'Yellow' }
             elseif ($item.Value -like '*difference*') { 'Yellow' }
             else { 'Red' }
    Write-Host ("  {0,-20} {1}" -f $item.Key, $item.Value) -ForegroundColor $color
}
Write-Host ""

#endregion
