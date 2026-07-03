<#
.SYNOPSIS
    Removes a security group exclusion from all Conditional Access policies in Entra ID.

.DESCRIPTION
    Discovers every Conditional Access policy in the tenant and removes the specified group
    from each policy's excluded groups (conditions.users.excludeGroups). Other existing
    exclusions are preserved. Policies that do not exclude the group are skipped.

    The group can be supplied by object ID or by display name.

.PARAMETER GroupId
    The object ID (GUID) of the group to remove from the exclusions.

.PARAMETER GroupName
    The display name of the group to remove. Resolved to an object ID via Graph.

.PARAMETER WhatIf
    Shows which policies would be updated without making any changes.

.EXAMPLE
    .\Remove-CAExclusionGroup.ps1 -GroupId 11111111-2222-3333-4444-555555555555

.EXAMPLE
    .\Remove-CAExclusionGroup.ps1 -GroupName "CA Break Glass Exclusions"

.EXAMPLE
    .\Remove-CAExclusionGroup.ps1 -GroupName "CA Break Glass Exclusions" -WhatIf

.NOTES
    Requires the Microsoft.Graph.Identity.SignIns and Microsoft.Graph.Groups PowerShell modules.
    The account used must have Policy.ReadWrite.ConditionalAccess and Group.Read.All permissions.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true, ParameterSetName = 'ById', HelpMessage = 'Object ID (GUID) of the group to remove from exclusions.')]
    [string]$GroupId,

    [Parameter(Mandatory = $true, ParameterSetName = 'ByName', HelpMessage = 'Display name of the group to remove from exclusions.')]
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
Write-Host "Exclusion group to remove: $targetGroupName ($targetGroupId)" -ForegroundColor Cyan

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

        if ($excludeGroups -notcontains $targetGroupId) {
            Write-Host "  Group not excluded. Skipping." -ForegroundColor Green
            $skippedCount++
            continue
        }

        $newExcludeGroups = @($excludeGroups | Where-Object { $_ -and $_ -ne $targetGroupId })

        $body = @{
            conditions = @{
                users = @{
                    excludeGroups = $newExcludeGroups
                }
            }
        }

        if ($PSCmdlet.ShouldProcess($policy.DisplayName, "Remove exclusion group '$targetGroupName'")) {
            Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id -BodyParameter $body -ErrorAction Stop
            Write-Host "  Exclusion removed." -ForegroundColor Green
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
Write-Host "  Policies updated        : $updatedCount" -ForegroundColor Green
Write-Host "  Not excluded (skipped)  : $skippedCount" -ForegroundColor DarkGray
if ($failCount -gt 0) {
    Write-Host "  Failures                : $failCount" -ForegroundColor Red
}
