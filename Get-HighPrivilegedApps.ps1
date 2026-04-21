<#
.SYNOPSIS
    Report all high-privileged enterprise applications in the tenant
.DESCRIPTION
    This script queries Entra ID for enterprise applications (service principals)
    and identifies those with high-privilege permissions or admin consent granted.
    Reports on apps with dangerous permissions and their consent status.
#>

# High-risk permissions to look for
$HighRiskPermissions = @(
    "Directory.ReadWrite.All"
    "RoleManagement.ReadWrite.Directory"
    "AppRoleAssignment.ReadWrite.All"
    "User.ReadWrite.All"
    "Group.ReadWrite.All"
    "Organization.Read.All"
    "MailboxSettings.ReadWrite"
    "Exchange.ManageAsApp"
    "Mail.Send"
    "Calendars.ReadWrite"
    "Contacts.ReadWrite"
    "Files.ReadWrite.All"
    "Sites.ReadWrite.All"
    "TeamSettings.ReadWrite.All"
    "Policy.ReadWrite.PermissionGrant"
)

Write-Host "Querying Entra ID for high-privileged enterprise applications..." -ForegroundColor Cyan

# Get all service principals with admin consent
$servicePrincipals = Get-MgServicePrincipal -All -PageSize 999 | Where-Object {
    $_.ServicePrincipalType -eq "Application"
}

$highPrivilegedApps = @()

foreach ($sp in $servicePrincipals) {
    $appRoles = $sp.AppRoles
    $oauth2PermissionScopes = $sp.Oauth2PermissionScopes
    
    # Check for high-risk app roles (application permissions)
    $riskyAppRoles = $appRoles | Where-Object { $_.Value -in $HighRiskPermissions }
    
    # Check for high-risk delegated permissions
    $riskyDelegatedPerms = $oauth2PermissionScopes | Where-Object { $_.Value -in $HighRiskPermissions }
    
    # Get consent grants for this app
    $consentGrants = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($sp.Id)'" -ErrorAction SilentlyContinue
    $adminConsented = $consentGrants | Where-Object { $_.ConsentType -eq "AllPrincipals" }
    
    # Add to report if has risky permissions or admin consent
    if ($riskyAppRoles -or $riskyDelegatedPerms -or $adminConsented) {
        $highPrivilegedApps += [PSCustomObject]@{
            DisplayName                 = $sp.DisplayName
            AppId                       = $sp.AppId
            ServicePrincipalId          = $sp.Id
            Publisher                   = $sp.PublisherName
            HighRiskApplicationPermissions = ($riskyAppRoles.Value -join ", ") -or "None"
            HighRiskDelegatedPermissions = ($riskyDelegatedPerms.Value -join ", ") -or "None"
            TotalAppRoles              = $appRoles.Count
            TotalDelegatedScopes       = $oauth2PermissionScopes.Count
            AdminConsentGranted        = if ($adminConsented) { "Yes" } else { "No" }
            CreatedDateTime            = $sp.CreatedDateTime
            ServicePrincipalType       = $sp.ServicePrincipalType
        }
    }
}

# Display results
Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "HIGH-PRIVILEGED ENTERPRISE APPLICATIONS" -ForegroundColor Yellow
Write-Host "========================================`n" -ForegroundColor Yellow

if ($highPrivilegedApps.Count -gt 0) {
    Write-Host "Found $($highPrivilegedApps.Count) high-privileged applications:`n" -ForegroundColor Green
    
    $highPrivilegedApps | Format-Table -AutoSize -Property `
        DisplayName, `
        Publisher, `
        AdminConsentGranted, `
        TotalAppRoles, `
        TotalDelegatedScopes
    
    Write-Host "`nDetailed Report:" -ForegroundColor Cyan
    $highPrivilegedApps | ForEach-Object {
        Write-Host "`n--- $($_.DisplayName) ---" -ForegroundColor Magenta
        Write-Host "  App ID: $($_.AppId)"
        Write-Host "  Service Principal ID: $($_.ServicePrincipalId)"
        Write-Host "  Publisher: $($_.Publisher)"
        Write-Host "  Admin Consent Granted: $($_.AdminConsentGranted)"
        Write-Host "  High-Risk App Permissions: $($_.HighRiskApplicationPermissions)"
        Write-Host "  High-Risk Delegated Permissions: $($_.HighRiskDelegatedPermissions)"
        Write-Host "  Total App Roles: $($_.TotalAppRoles)"
        Write-Host "  Total Delegated Scopes: $($_.TotalDelegatedScopes)"
        Write-Host "  Created: $($_.CreatedDateTime)"
    }
    
    # Export to CSV
    $csvPath = "C:\temp\GitHub\HighPrivilegedApps_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $highPrivilegedApps | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`n✓ Report exported to: $csvPath`n" -ForegroundColor Green
    
} else {
    Write-Host "No high-privileged applications with risky permissions found.`n" -ForegroundColor Green
}

Write-Host "Scan completed." -ForegroundColor Cyan
