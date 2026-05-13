#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Duplicates the Standard Preset Security Policy in Microsoft 365.

.DESCRIPTION
    Creates copies of all policies bundled within the Standard Preset Security Policy:
      - Anti-spam (HostedContentFilterPolicy)
      - Anti-malware (MalwareFilterPolicy)
      - Anti-phishing (AntiPhishPolicy)
      - Safe Links (SafeLinksPolicy)         [requires Defender for Office 365 Plan 1 or 2]
      - Safe Attachments (SafeAttachmentsPolicy) [requires Defender for Office 365 Plan 1 or 2]

    Each policy is created as a standalone custom policy and its associated rule is created
    in a DISABLED state so it does not immediately affect users. Recipient conditions are
    set on each rule and the priority is set to 0 so it appears immediately below the current
    preset policies. You must configure conditions (e.g. applies to domain) and enable the
    rules before they take effect.

.PARAMETER PolicyName
    The name for the duplicated policies. Defaults to "Standard Preset Policy - Duplicate".
    The same name is used for all five policy types and their corresponding rules.

.PARAMETER RecipientDomain
    The accepted domain to scope all rules to (e.g. contoso.com).
    If neither -RecipientDomain nor -RecipientGroup is specified, the tenant's
    default accepted domain is used automatically.

.PARAMETER RecipientGroup
    A mail-enabled security group or distribution group (display name or email)
    to scope all rules to (e.g. "All Staff" or allstaff@contoso.com).
    Mutually exclusive with -RecipientDomain.

.PARAMETER PolicyTypes
    One or more policy types to duplicate. Valid values:
      AntiSpam, AntiMalware, AntiPhish, SafeLinks, SafeAttachments, All
    If omitted, an interactive menu is displayed.

.EXAMPLE
    .\Duplicate-StandardPresetSecurityPolicy.ps1

.EXAMPLE
    .\Duplicate-StandardPresetSecurityPolicy.ps1 -PolicyName "Contoso Standard Policy"

.EXAMPLE
    .\Duplicate-StandardPresetSecurityPolicy.ps1 -RecipientDomain "contoso.com"

.EXAMPLE
    .\Duplicate-StandardPresetSecurityPolicy.ps1 -PolicyName "Contoso Standard Policy" -RecipientGroup "All Staff"

.EXAMPLE
    .\Duplicate-StandardPresetSecurityPolicy.ps1 -PolicyTypes AntiSpam, AntiPhish

.EXAMPLE
    .\Duplicate-StandardPresetSecurityPolicy.ps1 -PolicyTypes All

.NOTES
    Prerequisites:
    - ExchangeOnlineManagement module v3.0 or later must be installed:
        Install-Module -Name ExchangeOnlineManagement
    - You must have appropriate admin permissions in Microsoft 365
      (Security Administrator or Global Administrator)
    - Script connects to Exchange Online if not already connected

    IMPORTANT: The duplicate policies are created DISABLED with no recipient conditions.
    You must configure conditions (e.g. applies to domain) and enable the rules before
    they take effect.
#>

#Requires -Module ExchangeOnlineManagement

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $false)]
    [string]$PolicyName = "Standard Preset Policy - Duplicate",

    # Scope rules to a specific accepted domain (e.g. contoso.com).
    # If neither -RecipientDomain nor -RecipientGroup is supplied, the tenant's
    # default accepted domain is used automatically.
    [Parameter(Mandatory = $false)]
    [string]$RecipientDomain,

    # Scope rules to a mail-enabled security group or distribution group
    # (display name or email address, e.g. "All Staff" or allstaff@contoso.com).
    # Mutually exclusive with -RecipientDomain.
    [Parameter(Mandatory = $false)]
    [string]$RecipientGroup,

    # Which policy types to duplicate. If omitted, an interactive menu is shown.
    [Parameter(Mandatory = $false)]
    [ValidateSet('AntiSpam', 'AntiMalware', 'AntiPhish', 'SafeLinks', 'SafeAttachments', 'All')]
    [string[]]$PolicyTypes
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

# Builds a parameter hashtable from a policy object, stripping metadata and read-only properties.
function Get-CopyableParams {
    param (
        [Parameter(Mandatory)]
        [object]$Policy,

        [string[]]$ExcludeProperties = @(
            'ZapEnabled', 'EnableSuspiciousSafetyTip'
        )
    )

    $metadataProps = @(
        'Identity', 'Id', 'ExchangeObjectId', 'Guid', 'DistinguishedName', 'Name',
        'WhenCreated', 'WhenChanged', 'WhenCreatedUTC', 'WhenChangedUTC',
        'ObjectState', 'IsDefault', 'IsBuiltIn', 'IsPreset', 'RecommendedPolicyType',
        'ExchangeVersion', 'OrganizationId', 'PSShowComputerName', 'PSComputerName',
        'RunspaceId', 'AdminDisplayName', 'Rules', 'ObjectCategory', 'IsValid', 'ObjectClass',
        'OriginatingServer', 'OrganizationalUnitRoot', 'IsPolicyOverrideApplied', 'IsBuiltInProtection',
        'EnableOrganizationBranding', 'EnableBlockingEncryptedAttachments', 'QuarantineTagForBlockingEncryptedAttachments'
    )

    $skip = $metadataProps + $ExcludeProperties

    $params = @{}
    foreach ($prop in $Policy.PSObject.Properties) {
        if ($prop.Name -in $skip) { continue }
        if ($null -eq $prop.Value) { continue }
        # IntraOrgFilterState can only be set when non-default; skip when value is 'Default'
        if ($prop.Name -eq 'IntraOrgFilterState' -and $prop.Value -eq 'Default') { continue }
        # Skip empty arrays/collections — passing empty arrays to some -* params causes errors
        if ($prop.Value -is [System.Collections.ICollection] -and $prop.Value.Count -eq 0) { continue }
        $params[$prop.Name] = $prop.Value
    }

    return $params
}

#endregion

# Note: each policy type has its own tenant-specific timestamp suffix, so the preset name
# is discovered separately per policy type below.

# Build the rule condition hashtable.
# Priority: explicit -RecipientGroup > explicit -RecipientDomain > auto-detected default domain.
$ruleCondition = @{}

if ($RecipientGroup) {
    $ruleCondition['SentToMemberOf'] = $RecipientGroup
    $conditionDescription = "Group: $RecipientGroup"
}
elseif ($RecipientDomain) {
    $ruleCondition['RecipientDomainIs'] = $RecipientDomain
    $conditionDescription = "Domain: $RecipientDomain"
}
else {
    $defaultDomain = (Get-AcceptedDomain | Where-Object { $_.Default -eq $true } | Select-Object -First 1).DomainName
    if (-not $defaultDomain) {
        Write-Error "Could not determine the tenant's default accepted domain. Specify -RecipientDomain or -RecipientGroup."
        return
    }
    $ruleCondition['RecipientDomainIs'] = $defaultDomain
    $conditionDescription = "Domain: $defaultDomain (default, auto-detected)"
}

#region --- Policy Type Selection Menu ---

$validTypes = @('AntiSpam', 'AntiMalware', 'AntiPhish', 'SafeLinks', 'SafeAttachments')

if (-not $PolicyTypes) {
    Write-Host ""
    Write-Host "Which policy types do you want to duplicate?" -ForegroundColor Cyan
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
Write-Host "  Duplicating Standard Preset Security Policy" -ForegroundColor Cyan
Write-Host "  New policy name: '$PolicyName'" -ForegroundColor Cyan
Write-Host "  Rule condition : $conditionDescription" -ForegroundColor Cyan
Write-Host "  Policy types   : $($PolicyTypes -join ', ')" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan

#region --- Anti-Spam Policy ---

if ('AntiSpam' -in $PolicyTypes) {

Write-Host ""
Write-Host "[Anti-Spam] HostedContentFilterPolicy" -ForegroundColor Yellow

try {
    $sourcePresetName = (Get-HostedContentFilterPolicy |
        Where-Object { $_.Name -like 'Standard Preset Security Policy*' } |
        Select-Object -First 1).Name
    if (-not $sourcePresetName) { throw "Standard Preset Security Policy not found for Anti-Spam." }
    $source = Get-HostedContentFilterPolicy -Identity $sourcePresetName -ErrorAction Stop

    $existing = Get-HostedContentFilterPolicy -Identity $PolicyName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Warning "  Anti-Spam policy '$PolicyName' already exists — skipping."
        $results['AntiSpam'] = 'Skipped (already exists)'
    }
    else {
        if ($PSCmdlet.ShouldProcess($PolicyName, 'New-HostedContentFilterPolicy')) {
            $params = Get-CopyableParams -Policy $source
            New-HostedContentFilterPolicy -Name $PolicyName @params -ErrorAction Stop | Out-Null
            New-HostedContentFilterRule -Name $PolicyName -HostedContentFilterPolicy $PolicyName @ruleCondition `
                -Enabled $false -Priority 0 -ErrorAction Stop | Out-Null
            Write-Host "  Policy and rule created (rule disabled)." -ForegroundColor Green
            $results['AntiSpam'] = 'Created'
        }
    }
}
catch {
    Write-Warning "  Failed to duplicate Anti-Spam policy: $($_.Exception.Message)"
    $results['AntiSpam'] = "Failed: $($_.Exception.Message)"
}

} # end AntiSpam

#endregion

#region --- Anti-Malware Policy ---

if ('AntiMalware' -in $PolicyTypes) {

Write-Host ""
Write-Host "[Anti-Malware] MalwareFilterPolicy" -ForegroundColor Yellow

try {
    $sourcePresetName = (Get-MalwareFilterPolicy |
        Where-Object { $_.Name -like 'Standard Preset Security Policy*' } |
        Select-Object -First 1).Name
    if (-not $sourcePresetName) { throw "Standard Preset Security Policy not found for Anti-Malware." }
    $source = Get-MalwareFilterPolicy -Identity $sourcePresetName -ErrorAction Stop

    $existing = Get-MalwareFilterPolicy -Identity $PolicyName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Warning "  Anti-Malware policy '$PolicyName' already exists — skipping."
        $results['AntiMalware'] = 'Skipped (already exists)'
    }
    else {
        if ($PSCmdlet.ShouldProcess($PolicyName, 'New-MalwareFilterPolicy')) {
            $params = Get-CopyableParams -Policy $source
            New-MalwareFilterPolicy -Name $PolicyName @params -ErrorAction Stop | Out-Null
            New-MalwareFilterRule -Name $PolicyName -MalwareFilterPolicy $PolicyName @ruleCondition `
                -Enabled $false -Priority 0 -ErrorAction Stop | Out-Null
            Write-Host "  Policy and rule created (rule disabled)." -ForegroundColor Green
            $results['AntiMalware'] = 'Created'
        }
    }
}
catch {
    Write-Warning "  Failed to duplicate Anti-Malware policy: $($_.Exception.Message)"
    $results['AntiMalware'] = "Failed: $($_.Exception.Message)"
}

} # end AntiMalware

#endregion

#region --- Anti-Phishing Policy ---

if ('AntiPhish' -in $PolicyTypes) {

Write-Host ""
Write-Host "[Anti-Phishing] AntiPhishPolicy" -ForegroundColor Yellow

try {
    $sourcePresetName = (Get-AntiPhishPolicy |
        Where-Object { $_.Name -like 'Standard Preset Security Policy*' } |
        Select-Object -First 1).Name
    if (-not $sourcePresetName) { throw "Standard Preset Security Policy not found for Anti-Phishing." }
    $source = Get-AntiPhishPolicy -Identity $sourcePresetName -ErrorAction Stop

    $existing = Get-AntiPhishPolicy -Identity $PolicyName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Warning "  Anti-Phishing policy '$PolicyName' already exists — skipping."
        $results['AntiPhish'] = 'Skipped (already exists)'
    }
    else {
        if ($PSCmdlet.ShouldProcess($PolicyName, 'New-AntiPhishPolicy')) {
            $params = Get-CopyableParams -Policy $source
            New-AntiPhishPolicy -Name $PolicyName @params -ErrorAction Stop | Out-Null
            New-AntiPhishRule -Name $PolicyName -AntiPhishPolicy $PolicyName @ruleCondition `
                -Enabled $false -Priority 0 -ErrorAction Stop | Out-Null
            Write-Host "  Policy and rule created (rule disabled)." -ForegroundColor Green
            $results['AntiPhish'] = 'Created'
        }
    }
}
catch {
    Write-Warning "  Failed to duplicate Anti-Phishing policy: $($_.Exception.Message)"
    $results['AntiPhish'] = "Failed: $($_.Exception.Message)"
}

} # end AntiPhish

#endregion

#region --- Safe Links Policy ---

if ('SafeLinks' -in $PolicyTypes) {

Write-Host ""
Write-Host "[Safe Links] SafeLinksPolicy" -ForegroundColor Yellow

try {
    $sourcePresetName = (Get-SafeLinksPolicy |
        Where-Object { $_.Name -like 'Standard Preset Security Policy*' } |
        Select-Object -First 1).Name
    if (-not $sourcePresetName) { throw "Standard Preset Security Policy not found for Safe Links." }
    $source = Get-SafeLinksPolicy -Identity $sourcePresetName -ErrorAction Stop

    $existing = Get-SafeLinksPolicy -Identity $PolicyName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Warning "  Safe Links policy '$PolicyName' already exists — skipping."
        $results['SafeLinks'] = 'Skipped (already exists)'
    }
    else {
        if ($PSCmdlet.ShouldProcess($PolicyName, 'New-SafeLinksPolicy')) {
            $params = Get-CopyableParams -Policy $source
            New-SafeLinksPolicy -Name $PolicyName @params -ErrorAction Stop | Out-Null
            New-SafeLinksRule -Name $PolicyName -SafeLinksPolicy $PolicyName @ruleCondition `
                -Enabled $false -Priority 0 -ErrorAction Stop | Out-Null
            Write-Host "  Policy and rule created (rule disabled)." -ForegroundColor Green
            $results['SafeLinks'] = 'Created'
        }
    }
}
catch {
    $msg = $_.Exception.Message
    if ($msg -match 'recognized|licensed|subscription|not found') {
        Write-Warning "  Safe Links not available — requires Defender for Office 365 Plan 1 or 2."
        $results['SafeLinks'] = 'Skipped (license not available)'
    }
    else {
        Write-Warning "  Failed to duplicate Safe Links policy: $msg"
        $results['SafeLinks'] = "Failed: $msg"
    }
}

} # end SafeLinks

#endregion

#region --- Safe Attachments Policy ---

if ('SafeAttachments' -in $PolicyTypes) {

Write-Host ""
Write-Host "[Safe Attachments] SafeAttachmentsPolicy" -ForegroundColor Yellow

try {
    $sourcePresetName = (Get-SafeAttachmentPolicy |
        Where-Object { $_.Name -like 'Standard Preset Security Policy*' } |
        Select-Object -First 1).Name
    if (-not $sourcePresetName) { throw "Standard Preset Security Policy not found for Safe Attachments." }
    $source = Get-SafeAttachmentPolicy -Identity $sourcePresetName -ErrorAction Stop

    $existing = Get-SafeAttachmentPolicy -Identity $PolicyName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Warning "  Safe Attachments policy '$PolicyName' already exists — skipping."
        $results['SafeAttachments'] = 'Skipped (already exists)'
    }
    else {
        if ($PSCmdlet.ShouldProcess($PolicyName, 'New-SafeAttachmentPolicy')) {
            $params = Get-CopyableParams -Policy $source
            New-SafeAttachmentPolicy -Name $PolicyName @params -ErrorAction Stop | Out-Null
            New-SafeAttachmentRule -Name $PolicyName -SafeAttachmentPolicy $PolicyName @ruleCondition `
                -Enabled $false -Priority 0 -ErrorAction Stop | Out-Null
            Write-Host "  Policy and rule created (rule disabled)." -ForegroundColor Green
            $results['SafeAttachments'] = 'Created'
        }
    }
}
catch {
    $msg = $_.Exception.Message
    if ($msg -match 'recognized|licensed|subscription|not found') {
        Write-Warning "  Safe Attachments not available — requires Defender for Office 365 Plan 1 or 2."
        $results['SafeAttachments'] = 'Skipped (license not available)'
    }
    else {
        Write-Warning "  Failed to duplicate Safe Attachments policy: $msg"
        $results['SafeAttachments'] = "Failed: $msg"
    }
}

} # end SafeAttachments

#endregion

#region --- Summary ---

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  Summary" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
foreach ($item in $results.GetEnumerator()) {
    $color = if ($item.Value -eq 'Created') { 'Green' }
             elseif ($item.Value -like 'Skipped*') { 'Yellow' }
             else { 'Red' }
    Write-Host ("  {0,-20} {1}" -f $item.Key, $item.Value) -ForegroundColor $color
}
Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor White
Write-Host "  1. In the Microsoft Defender portal, locate the new custom policies named '$PolicyName'."
Write-Host "  2. Each rule has been scoped to: $conditionDescription"
Write-Host "     Adjust recipient conditions if you need a different scope."
Write-Host "  3. Enable each rule when ready to apply protection."
Write-Host ""

#endregion
