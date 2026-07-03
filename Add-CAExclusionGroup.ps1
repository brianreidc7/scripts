<#
.SYNOPSIS
    Adds a security group as an exclusion to all Conditional Access policies in Entra ID.

.DESCRIPTION
    Discovers every Conditional Access policy in the tenant and adds the specified group
    to each policy's excluded groups (conditions.users.excludeGroups). Existing exclusions
    are preserved. Policies that already exclude the group are skipped.

    The group can be supplied by object ID or by display name.

.PARAMETER GroupId
    The object ID (GUID) of the group to add as an exclusion.

.PARAMETER GroupName
    The display name of the group to add as an exclusion. Resolved to an object ID via Graph.

.PARAMETER WhatIf
    Shows which policies would be updated without making any changes.

.EXAMPLE
    .\Add-FreepointCAExclusionGroup.ps1 -GroupId 11111111-2222-3333-4444-555555555555

.EXAMPLE
    .\Add-FreepointCAExclusionGroup.ps1 -GroupName "CA Break Glass Exclusions"

.EXAMPLE
    .\Add-FreepointCAExclusionGroup.ps1 -GroupName "CA Break Glass Exclusions" -WhatIf

.NOTES
    Requires the Microsoft.Graph.Identity.SignIns and Microsoft.Graph.Groups PowerShell modules.
    The account used must have Policy.ReadWrite.ConditionalAccess and Group.Read.All permissions.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true, ParameterSetName = 'ById', HelpMessage = 'Object ID (GUID) of the group to exclude.')]
    [string]$GroupId,

    [Parameter(Mandatory = $true, ParameterSetName = 'ByName', HelpMessage = 'Display name of the group to exclude.')]
    [string]$GroupName
)

Import-Module Microsoft.Graph.Identity.SignIns
Import-Module Microsoft.Graph.Groups

# Connect to Microsoft Graph only if not already connected
$mgContext = Get-MgContext
if (-not $mgContext) {
    Connect-MgGraph -TenantID 8b7b4f70-e59d-4f45-8a53-b834072c17ad -Scopes Policy.ReadWrite.ConditionalAccess,Group.Read.All -NoWelcome
}
else {
    Write-Host "Already connected to Microsoft Graph." -ForegroundColor DarkGray
}

# Resolve the group to an object ID and display name
try {
    if ($PSCmdlet.ParameterSetName -eq 'ByName') {
        $group = Get-MgGroup -Filter "displayName eq '$GroupName'" -All -ErrorAction Stop
        if (-not $group) {
            Write-Error "No group found with display name '$GroupName'."
            return
        }
        if ($group.Count -gt 1) {
            Write-Error "Multiple groups found with display name '$GroupName'. Use -GroupId to specify the exact group."
            return
        }
    }
    else {
        $group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
    }
}
catch {
    Write-Error "Failed to resolve group: $_"
    return
}

$targetGroupId   = $group.Id
$targetGroupName = $group.DisplayName
Write-Host "Exclusion group: $targetGroupName ($targetGroupId)" -ForegroundColor Cyan

# Discover all Conditional Access policies
try {
    $policies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
}
catch {
    Write-Error "Failed to retrieve Conditional Access policies: $_"
    return
}

if (-not $policies) {
    Write-Warning "No Conditional Access policies were found in the tenant."
    return
}

Write-Host "Discovered $($policies.Count) Conditional Access policies." -ForegroundColor Cyan

$updatedCount = 0
$skippedCount = 0
$failCount    = 0

foreach ($policy in $policies) {

    Write-Host "Processing: $($policy.DisplayName)" -ForegroundColor Cyan

    try {
        $excludeGroups = @($policy.Conditions.Users.ExcludeGroups)

        if ($excludeGroups -contains $targetGroupId) {
            Write-Host "  Already excluded. Skipping." -ForegroundColor Green
            $skippedCount++
            continue
        }

        $newExcludeGroups = @($excludeGroups + $targetGroupId | Where-Object { $_ })

        $body = @{
            conditions = @{
                users = @{
                    excludeGroups = $newExcludeGroups
                }
            }
        }

        if ($PSCmdlet.ShouldProcess($policy.DisplayName, "Add exclusion group '$targetGroupName'")) {
            Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id -BodyParameter $body -ErrorAction Stop
            Write-Host "  Exclusion added." -ForegroundColor Green
            $updatedCount++
        }
    }
    catch {
        Write-Warning "  Failed to update '$($policy.DisplayName)': $_"
        $failCount++
    }
}

Write-Host ""
Write-Host "--- Summary ---" -ForegroundColor White
Write-Host "  Policies updated       : $updatedCount" -ForegroundColor Green
Write-Host "  Already excluded (skip): $skippedCount" -ForegroundColor DarkGray
if ($failCount -gt 0) {
    Write-Host "  Failures               : $failCount" -ForegroundColor Red
}
