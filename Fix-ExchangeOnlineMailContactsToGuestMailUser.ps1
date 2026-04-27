#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Manages mail contacts and guest mail users in Exchange Online with three operational sections.

.DESCRIPTION
    This script provides three sections for managing mail contacts and guest mail users:
    
    Section 1: Export mail contacts from a domain and save identity attributes to CSV
    Section 2: Delete mail contacts that were previously exported to a CSV file
    Section 3: Update guest mail users by adding LegacyExchangeDN as X500 email addresses

.PARAMETER Section
    Optional. Specify which section to run (1, 2, or 3). If not specified, displays an interactive menu.

.PARAMETER DomainName
    Required for Section 1. The domain name to filter mail contacts by (e.g., 'contoso.com').

.PARAMETER OutputPath
    Optional for Section 1. The path to save the CSV file. If not specified, creates a timestamped file.

.PARAMETER InputPath
    Required for Sections 2 and 3. The path to the CSV file containing mail contact data.

.PARAMETER ExportToCSV
    Optional for Section 1. Switch to export results to CSV file.

.EXAMPLE
    # Interactive menu
    .\Fix-ExchangeOnlineMailContactsToGuestMailUser.ps1

.EXAMPLE
    # Section 1: Export mail contacts
    .\Fix-ExchangeOnlineMailContactsToGuestMailUser.ps1 -Section 1 -DomainName contoso.com -ExportToCSV

.EXAMPLE
    # Section 2: Delete exported mail contacts
    .\Fix-ExchangeOnlineMailContactsToGuestMailUser.ps1 -Section 2 -InputPath "contacts.csv"

.EXAMPLE
    # Section 3: Update guest mail users
    .\Fix-ExchangeOnlineMailContactsToGuestMailUser.ps1 -Section 3 -InputPath "contacts.csv"

.NOTES
    Prerequisites:
    - Must have ExchangeOnlineManagement module installed: Install-Module -Name ExchangeOnlineManagement
    - Must have appropriate permissions in Exchange Online
    - Must be connected to Exchange Online (script will prompt to connect if not already connected)

#>

param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("1", "2", "3")]
    [string]$Section,

    [Parameter(Mandatory = $false)]
    [string]$DomainName,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$InputPath,

    [Parameter(Mandatory = $false)]
    [bool]$ExportToCSV = $true
)

#Requires -Module ExchangeOnlineManagement

# Ensure we're connected to Exchange Online
try {
    $testConnection = Get-OrganizationConfig -ErrorAction Stop
    Write-Host "✓ Connected to Exchange Online" -ForegroundColor Green
}
catch {
    Write-Host "⚠ Not connected to Exchange Online. Initiating connection..." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowProgress $true
}

# Display menu if no section specified
if ([string]::IsNullOrEmpty($Section)) {
    Write-Host "`n" -ForegroundColor Cyan
    Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║     Mail Contact Management - Interactive Menu             ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  1  - Export mail contacts from domain to CSV" -ForegroundColor Green
    Write-Host "  2  - Delete exported mail contacts" -ForegroundColor Yellow
    Write-Host "  3  - Update guest mail users with LegacyExchangeDN" -ForegroundColor Blue
    Write-Host "  0  - Exit" -ForegroundColor Gray
    Write-Host ""
    Write-Host "════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    
    $Section = Read-Host "`nEnter your choice (0-3)"
    
    if ($Section -eq "0") {
        Write-Host "Exiting..." -ForegroundColor Cyan
        exit 0
    }
    
    if ($Section -notin @("1", "2", "3")) {
        Write-Host "✗ Invalid choice. Please run the script again." -ForegroundColor Red
        exit 1
    }
}

# ============================================================================
# SECTION 1: Export Mail Contacts
# ============================================================================
if ($Section -eq "1") {
    Write-Host "`n" -ForegroundColor Cyan
    Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║  Section 1: Export Mail Contacts from Domain               ║" -ForegroundColor Green
    Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    
    if ([string]::IsNullOrEmpty($DomainName)) {
        $DomainName = Read-Host "`nEnter the domain name (e.g., contoso.com)"
    }
    
    if ([string]::IsNullOrEmpty($DomainName)) {
        Write-Host "⚠ Domain name is required" -ForegroundColor Red
        exit 1
    }
    
    try {
        $DomainName = $DomainName.ToLower().Trim()
        Write-Host "`nExtracting mail contacts for domain: $DomainName" -ForegroundColor Cyan

        Write-Host "Retrieving mail contacts..." -ForegroundColor Cyan
        $mailcontacts = Get-MailContact -Filter "EmailAddresses -like '*@$DomainName'" -ResultSize Unlimited -ErrorAction Stop

        if ($null -eq $mailcontacts -or $mailcontacts.Count -eq 0) {
            Write-Host "⚠ No mail contacts found for domain: $DomainName" -ForegroundColor Yellow
            exit 0
        }

        Write-Host "Found $($mailcontacts.Count) mail contact(s)" -ForegroundColor Green

        Write-Host "Retrieving directory groups and members for lookup..." -ForegroundColor Cyan
        $groupMemberMap = @{}
        try {
            $allGroups = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop
            foreach ($group in $allGroups) {
                try {
                    $members = Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue
                    if ($null -ne $members) {
                        if ($members -is [array]) {
                            foreach ($member in $members) {
                                if (-not $groupMemberMap.ContainsKey($member.DistinguishedName)) {
                                    $groupMemberMap[$member.DistinguishedName] = @()
                                }
                                $groupMemberMap[$member.DistinguishedName] += $group.PrimarySmtpAddress
                            }
                        }
                        else {
                            if (-not $groupMemberMap.ContainsKey($members.DistinguishedName)) {
                                $groupMemberMap[$members.DistinguishedName] = @()
                            }
                            $groupMemberMap[$members.DistinguishedName] += $group.PrimarySmtpAddress
                        }
                    }
                }
                catch {
                    # Continue if unable to retrieve members for this group
                }
            }
            Write-Host "✓ Loaded $($groupMemberMap.Count) members from distribution groups" -ForegroundColor Green
        }
        catch {
            Write-Host "⚠ Could not retrieve groups and members: $_" -ForegroundColor Yellow
        }

        Write-Host "Extracting attributes..." -ForegroundColor Cyan
        $results = @()

        foreach ($mailcontact in $mailcontacts) {
            $mcDetails = Get-MailContact -Identity $mailcontact.Identity -ErrorAction SilentlyContinue

            if ($null -ne $mcDetails) {
                # Look up groups this mail contact is a member of
                $memberOfGroups = @()
                if ($groupMemberMap.ContainsKey($mcDetails.DistinguishedName)) {
                    $memberOfGroups = @($groupMemberMap[$mcDetails.DistinguishedName])
                }
                
                $result = [PSCustomObject]@{
                    DisplayName              = $mailcontact.DisplayName
                    UserPrincipalName        = $mailcontact.UserPrincipalName
                    PrimarySmtpAddress       = $mailcontact.PrimarySmtpAddress
                    RecipientType            = $mailcontact.RecipientType
                    RecipientTypeDetails     = $mailcontact.RecipientTypeDetails
                    LegacyExchangeDN         = $mcDetails.LegacyExchangeDN
                    EmailAddresses           = $mcDetails.EmailAddresses -join '; '
                    Alias                    = $mailcontact.Alias
                    OrganizationalUnit       = $mailcontact.OrganizationalUnit
                    Identity                 = $mailcontact.Identity
                    ExternalDirectoryObjectId = $mailcontact.ExternalDirectoryObjectId
                    CreatedDate              = $mcDetails.WhenCreatedUTC
                    WhenChangedUTC           = $mcDetails.WhenChangedUTC
                    MemberOfGroups           = $memberOfGroups -join '; '
                }
                $results += $result
            }
        }

        if ($results.Count -gt 0) {
            Write-Host "`n✓ Successfully extracted attributes for $($results.Count) recipient(s)" -ForegroundColor Green
            Write-Host "`nResults:`n" -ForegroundColor Cyan
            $results | Format-Table -AutoSize

            if ($ExportToCSV -or -not [string]::IsNullOrEmpty($OutputPath)) {
                if ([string]::IsNullOrEmpty($OutputPath)) {
                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    $OutputPath = "ExchangeOnlineContacts_${DomainName}_${timestamp}.csv"
                }

                $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
                Write-Host "`n✓ Results exported to: $OutputPath" -ForegroundColor Green
            }
        }
        else {
            Write-Host "⚠ No valid mail contacts could be retrieved" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "✗ Error: $_" -ForegroundColor Red
        exit 1
    }
}

# ============================================================================
# SECTION 2: Delete Mail Contacts
# ============================================================================
elseif ($Section -eq "2") {
    Write-Host "`n" -ForegroundColor Cyan
    Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "║  Section 2: Delete Exported Mail Contacts                  ║" -ForegroundColor Yellow
    Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
    
    if ([string]::IsNullOrEmpty($InputPath)) {
        $InputPath = Read-Host "`nEnter the path to the CSV file"
    }
    
    if (-not (Test-Path $InputPath)) {
        Write-Host "✗ CSV file not found: $InputPath" -ForegroundColor Red
        exit 1
    }
    
    try {
        $mailcontacts = Import-Csv -Path $InputPath
        
        if ($null -eq $mailcontacts -or $mailcontacts.Count -eq 0) {
            Write-Host "⚠ No mail contacts found in CSV file" -ForegroundColor Yellow
            exit 0
        }
        
        Write-Host "`nFound $($mailcontacts.Count) mail contact(s) to delete" -ForegroundColor Cyan
        Write-Host "`nMail contacts to be deleted:" -ForegroundColor Cyan
        $mailcontacts | Format-Table -Property DisplayName, PrimarySmtpAddress -AutoSize
        
        $confirmation = Read-Host "`nAre you sure you want to delete these contacts? (yes/no) [default: no]"
        if ($confirmation -notmatch '^y(es)?$') {
            Write-Host "Operation cancelled." -ForegroundColor Cyan
            exit 0
        }
        
        # Ensure Microsoft Graph connection for guest invitations
        $graphConnected = $false
        try {
            if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Identity.SignIns)) {
                Write-Host "⚠ Microsoft.Graph.Identity.SignIns module not installed. Guest invitations will be skipped." -ForegroundColor Yellow
                Write-Host "   Install with: Install-Module Microsoft.Graph.Identity.SignIns" -ForegroundColor Yellow
            }
            else {
                Import-Module Microsoft.Graph.Identity.SignIns -ErrorAction Stop
                $mgContext = Get-MgContext -ErrorAction SilentlyContinue
                if ($null -eq $mgContext) {
                    Write-Host "Connecting to Microsoft Graph (User.Invite.All scope)..." -ForegroundColor Cyan
                    Connect-MgGraph -Scopes "User.Invite.All", "User.ReadWrite.All" -NoWelcome -ErrorAction Stop
                }
                $graphConnected = $true
                Write-Host "✓ Connected to Microsoft Graph" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "⚠ Could not connect to Microsoft Graph: $_" -ForegroundColor Yellow
            Write-Host "   Guest invitations will be skipped." -ForegroundColor Yellow
        }

        $deleted = 0
        $failed = 0
        $invited = 0
        $inviteFailed = 0
        
        foreach ($contact in $mailcontacts) {
            try {
                Remove-MailContact -Identity $contact.Identity -Confirm:$false -ErrorAction Stop
                Write-Host "✓ Deleted: $($contact.DisplayName) ($($contact.PrimarySmtpAddress))" -ForegroundColor Green
                $deleted++

                # Invite the same email address as a guest, without sending an invitation
                if ($graphConnected) {
                    try {
                        $invitedEmail = $contact.PrimarySmtpAddress
                        $invitedName = if ([string]::IsNullOrWhiteSpace($contact.DisplayName)) { $invitedEmail } else { $contact.DisplayName }

                        $invitationParams = @{
                            InvitedUserEmailAddress = $invitedEmail
                            InvitedUserDisplayName  = $invitedName
                            InviteRedirectUrl       = "https://myapps.microsoft.com"
                            SendInvitationMessage   = $false
                        }

                        $invitation = New-MgInvitation @invitationParams -ErrorAction Stop
                        Write-Host "  ✓ Invited as guest: $invitedEmail (no invitation email sent)" -ForegroundColor Green
                        $invited++
                    }
                    catch {
                        Write-Host "  ✗ Failed to invite $($contact.PrimarySmtpAddress) as guest: $_" -ForegroundColor Red
                        $inviteFailed++
                    }
                }
            }
            catch {
                Write-Host "✗ Failed to delete $($contact.DisplayName): $_" -ForegroundColor Red
                $failed++
            }
        }
        
        Write-Host "`n" -ForegroundColor Cyan
        Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "║  Deletion Summary                                          ║" -ForegroundColor Yellow
        Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host "Successfully deleted: $deleted" -ForegroundColor Green
        Write-Host "Failed to delete: $failed" -ForegroundColor Red
        Write-Host "Guest invitations created: $invited" -ForegroundColor Green
        Write-Host "Guest invitations failed: $inviteFailed" -ForegroundColor Red
    }
    catch {
        Write-Host "✗ Error: $_" -ForegroundColor Red
        exit 1
    }
}

# ============================================================================
# SECTION 3: Update Guest Mail Users
# ============================================================================
elseif ($Section -eq "3") {
    Write-Host "`n" -ForegroundColor Cyan
    Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Blue
    Write-Host "║  Section 3: Update Guest Mail Users with LegacyExchangeDN  ║" -ForegroundColor Blue
    Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Blue
    
    if ([string]::IsNullOrEmpty($InputPath)) {
        $InputPath = Read-Host "`nEnter the path to the CSV file"
    }
    
    if (-not (Test-Path $InputPath)) {
        Write-Host "✗ CSV file not found: $InputPath" -ForegroundColor Red
        exit 1
    }
    
    try {
        $csvdata = Import-Csv -Path $InputPath
        
        if ($null -eq $csvdata -or $csvdata.Count -eq 0) {
            Write-Host "⚠ No data found in CSV file" -ForegroundColor Yellow
            exit 0
        }
        
        Write-Host "`nFound $($csvdata.Count) record(s) to process" -ForegroundColor Cyan
        
        $updated = 0
        $failed = 0
        $notfound = 0
        $skipped = 0
        
        foreach ($record in $csvdata) {
            try {
                $smtpAddress = $record.PrimarySmtpAddress
                Write-Host "`nProcessing: $smtpAddress" -ForegroundColor Cyan
                
                # Find guest user by SMTP address
                $guestUser = Get-MailUser -Filter "PrimarySmtpAddress -eq '$smtpAddress'" -ResultSize 1 -ErrorAction Stop
                
                if ($null -eq $guestUser) {
                    Write-Host "  ⚠ Guest user not found" -ForegroundColor Yellow
                    $notfound++
                    continue
                }
                
                # Create X500 address from LegacyExchangeDN
                $legacyDN = $record.LegacyExchangeDN
                
                if ([string]::IsNullOrEmpty($legacyDN)) {
                    Write-Host "  ⚠ LegacyExchangeDN is empty, skipping" -ForegroundColor Yellow
                    $skipped++
                    continue
                }
                
                $x500Address = "x500:$legacyDN"
                $currentAddresses = @($guestUser.EmailAddresses)
                
                # Check if X500 address already exists
                if ($currentAddresses -contains $x500Address) {
                    Write-Host "  ℹ X500 address already exists" -ForegroundColor Gray
                    $skipped++
                }
                else {
                    # Add X500 address
                    $newAddresses = $currentAddresses + $x500Address
                    Set-MailUser -Identity $guestUser.Identity -EmailAddresses $newAddresses -ErrorAction Stop
                    Write-Host "  ✓ Updated: Added X500 address $x500Address" -ForegroundColor Green
                    $updated++
                }
                
                # Add guest user to groups from MemberOfGroups column
                if (-not [string]::IsNullOrEmpty($record.MemberOfGroups)) {
                    $groupEmails = @($record.MemberOfGroups -split '; ' | Where-Object { $_ -match '\S' })
                    
                    foreach ($groupEmail in $groupEmails) {
                        try {
                            $group = Get-DistributionGroup -Filter "PrimarySmtpAddress -eq '$groupEmail'" -ErrorAction SilentlyContinue
                            
                            if ($null -eq $group) {
                                Write-Host "  ⚠ Group not found: $groupEmail" -ForegroundColor Yellow
                                continue
                            }
                            
                            # Check if user is already a member
                            $isMember = Get-DistributionGroupMember -Identity $group.Identity -ErrorAction SilentlyContinue | 
                                Where-Object { $_.PrimarySmtpAddress -eq $smtpAddress }
                            
                            if ($null -ne $isMember) {
                                Write-Host "  ℹ Already member of: $groupEmail" -ForegroundColor Gray
                            }
                            else {
                                Add-DistributionGroupMember -Identity $group.Identity -Member $guestUser.Identity -Confirm:$false -BypassSecurityGroupManagerCheck -ErrorAction Stop
                                Write-Host "  ✓ Added to group: $groupEmail" -ForegroundColor Green
                            }
                        }
                        catch {
                            Write-Host "  ✗ Failed to add to group $groupEmail : $_" -ForegroundColor Red
                        }
                    }
                }
            }
            catch {
                Write-Host "  ✗ Error: $_" -ForegroundColor Red
                $failed++
            }
        }
        
        Write-Host "`n" -ForegroundColor Cyan
        Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Blue
        Write-Host "║  Update Summary                                            ║" -ForegroundColor Blue
        Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Blue
        Write-Host "Successfully updated: $updated" -ForegroundColor Green
        Write-Host "Guest users not found: $notfound" -ForegroundColor Yellow
        Write-Host "Skipped (already exists or empty LegacyExchangeDN): $skipped" -ForegroundColor Gray
        Write-Host "Failed: $failed" -ForegroundColor Red
    }
    catch {
        Write-Host "✗ Error: $_" -ForegroundColor Red
        exit 1
    }
}

else {
    Write-Host "✗ Invalid section selected" -ForegroundColor Red
    exit 1
}

Write-Host "`n✓ Script completed." -ForegroundColor Cyan
