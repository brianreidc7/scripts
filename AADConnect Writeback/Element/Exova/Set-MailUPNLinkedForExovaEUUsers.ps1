cd C:\ExovaRebrand
$remoteEUdotLocal = Get-Credential -Message "Exova-EU Migration Account (Exova-EU.local)" -UserName "exova-eu\Brian.Reid"
$ExovaEUdotLocalDC = "eudc-ad-03.exova-eu.local"
$ElementEUDC = "usnjdcs040.eu.sys.element.com"
$ElementUSDC = "usnjdcs030.us.sys.element.com"
$ElementMEDC = "usnjdcs050.me.sys.element.com"
$ElementAPDC = "usnjdcs060.ap.sys.element.com"
$ElementCNDC = "usnjdcs070.cn.sys.element.com"
$ExovaEUDC = "eudc-gad-01.exova.com"
$ExovaNADC = "dc-ad02.bcamericas.com"

# Done
# Move all disabled accounts 
# Search-ADAccount -UsersOnly -AccountDisabled -SearchBase "ou=enabled,ou=staging,ou=accounts,dc=eu,dc=sys,dc=element,dc=com" -Server $ElementEUDC | Move-ADObject -TargetPath "ou=staging,ou=accounts,dc=eu,dc=sys,dc=element,dc=com" -Server $ElementEUDC

# Link mailboxes
# Updating linked account disables the source mailbox and sets the source UPN to come from exova-eu
# Therefore care when doing intentionally enabled accounts (re-enable at end)
# Link users in small batches and sync them to cloud. If disabled in Element can do in bigger batches. If enabled ensure remain enabled and treat with care
# Dont need to do all users when iCIMS/Cornerstone is fixed - only cannot start until that time.
    # Link | Enable if needed | Move object into sync in exova and exova-eu and sync then and resolve errors and repeat (2 days work)
$Users = Get-User -OrganizationalUnit "eu.sys.element.com/Accounts/Staging/"
ForEach ($User in $Users) {
    $LinkedAccount = "EXOVA-EU\" + $User.SamAccountName
    Write-Host $LinkedAccount
    Set-User $user -LinkedMasterAccount $LinkedAccount -LinkedDomainController $ExovaEUdotLocalDC -LinkedCredential $remoteEUdotLocal -DomainController $ElementEUDC
    # Where user is enabled, Element SamAccountName is changed so not same as exova-eu. Exova-EU is likely to be first.last
        # For users that are successfully linked with SamAccountName there is no issue in repeating them with first.last as they just error and say they are already linked
        $LinkedAccount = "EXOVA-EU\" + $User.firstname + "." + $user.lastname
        Set-User $user -LinkedMasterAccount $LinkedAccount -LinkedDomainController $ExovaEUdotLocalDC -LinkedCredential $remoteEUdotLocal -DomainController $ElementEUDC
    # Finally - some of these users are in the US or ME and so need to be processed in those domains instead of EU (last cmdlet)
    }

# Before run above - test with sync for the two already done users. Add new OU to AADConnect


# Export primary email address, samaccountname from Element
# Look for above in Exova-EU and set mail and UPN to primary



# Exova.com should be done already (policy driven in Exchange - might be a few exceptions) - there are lots of exceptions!

$Users = Get-MailUser -OrganizationalUnit "eu.sys.element.com/Accounts/Exova-Staging-Cloud/"
ForEach ($User in $Users) {
    $User | Select SamAccountName,PrimarySMTPAddress,DisplayName,UserPrincipalName | Export-CSV MailUserInElementStagingCloud.csv -Append -NoClobber -NoTypeInformation
}

# Need to update UPN in Element to match primary email address (have a script for this)
# No one in staging now - all done
# Lots in exova-staging-cloud (as has impact on cloud login) - not done
...


# Need to update users (archive first) in Exova-EU with mail and UPN from above CSV
    # Read $OldMail and $OldUPN from given user in Exova-EU
    # Get-Recipient in Element (all from same script - so needs to run on Exchange 2013 and connect to Exova-EU)
    # Determine $NewMailUPN
    # Write Mail and UPN to new value in Exova-EU (optional)
    # Write value to new spreadsheet so can add it to list for Carol

    #iCIMS Spreadsheet

    Set-ADServerSettings -ViewEntireForest $true
    cd C:\ExovaRebrand
    $ElementEUDC = "usnjdcs040.eu.sys.element.com"
    $ExovaEUDC = "eudc-gad-01.exova.com"
    $ExovaEUdotLocalDC = "eudc-ad-03.exova-eu.local"

    $SourceUsers = Import-CSV iCIMS-ExovaOnly-Source.csv

    $resultsarray =@()

    foreach ($SourceUser in $SourceUsers) {

        $TargetUser = Get-Recipient $SourceUser.Login

        $NewMailUPN = $targetUser.PrimarySmtpAddress.Address

        # Create a new custom object to hold our result.
        $userObject = new-object PSObject

        # Add our data to $contactObject as attributes using the add-member commandlet
        $userObject | add-member -membertype NoteProperty -name "SystemID" -Value $SourceUser.SystemID
        $userObject | add-member -membertype NoteProperty -name "Login" -Value $SourceUser.Login
        $userObject | add-member -membertype NoteProperty -name "OldPrimaryEmail" -Value $SourceUser.PrimaryEmail
        $userObject | add-member -membertype NoteProperty -name "ElementDN" -Value $TargetUserEx0v@123#
        $userObject | add-member -membertype NoteProperty -name "NewEmailAndUsername" -Value $NewMailUPN

        # Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
        $resultsarray += $userObject

        # Set-ADUser $SourceUser -Mail $NewMailUPN -UserPrincipalName $NewMailUPN -Server $ExovaEUdotLocalDC

    }
    
    # Export to CSV 
    $resultsarray| Export-csv iCIMS-ExovaOnly-Results.csv -notypeinformation

    #Cornerstone Spreadsheet (new source and new column names)
    # Filters applied to remove data before processing:
        # Remove if not containing @ in username
        # Include @exova.com, @warringtonfire.com and @bmtrada.com if listed
        # Include @isibfire.be 
        # First run of this filters out any row of data that does not have a recipient in Element

    Set-ADServerSettings -ViewEntireForest $true
    cd C:\ExovaRebrand
    $ElementEUDC = "usnjdcs040.eu.sys.element.com"
    $ExovaEUDC = "eudc-gad-01.exova.com"
    $ExovaEUdotLocalDC = "eudc-ad-03.exova-eu.local"

    $SourceUsers = Import-CSV Cornerstone-ExovaOnly-Source.csv

    $resultsarray =@()

    foreach ($SourceUser in $SourceUsers) {

        $TargetUser = Get-Recipient $SourceUser.Username

        $NewMailUPN = $targetUser.PrimarySmtpAddress.Address

        # Create a new custom object to hold our result.
        $userObject = new-object PSObject

        # Add our data to $contactObject as attributes using the add-member commandlet
        $userObject | add-member -membertype NoteProperty -name "UserID" -Value $SourceUser.UserID
        $userObject | add-member -membertype NoteProperty -name "Username" -Value $SourceUser.Username
        $userObject | add-member -membertype NoteProperty -name "OldUserEmail" -Value $SourceUser.UserEmail
        $userObject | add-member -membertype NoteProperty -name "UserFirstName" -Value $SourceUser.UserFirstName
        $userObject | add-member -membertype NoteProperty -name "UserLastName" -Value $SourceUser.UserLastName
        $userObject | add-member -membertype NoteProperty -name "LocalSystemID" -Value $SourceUser.LocalSystemID
        $userObject | add-member -membertype NoteProperty -name "UserLastAccess" -Value $SourceUser.UserLastAccess
        $userObject | add-member -membertype NoteProperty -name "RegionID" -Value $SourceUser.RegionID
        $userObject | add-member -membertype NoteProperty -name "ElementDN" -Value $TargetUser
        $userObject | add-member -membertype NoteProperty -name "NewEmailAndUsername" -Value $NewMailUPN

        # Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
        $resultsarray += $userObject

    }
    
    # Export to CSV 
    $resultsarray| Export-csv Cornerstone-ExovaOnly-Source-Round2.csv -notypeinformation


    # iCIMS Data

    ## Get DN of user in Exova-EU and build spreadsheet
    $SourceUsers = Import-CSV iCIMS-ExovaOnly-Source-Round2.csv

    $resultsarray =@()

    foreach ($SourceUser in $SourceUsers) {

        $ExovaUser = $SourceUser.Login  # Needed to use .Login (in the main) and  .OldPrimaryEmail and .NewEmailAndUsername for some users to find them as login name in app is wrong
        $TargetUser = Get-ADUser -Property mail -Filter {mail -eq $ExovaUser } -Server $ExovaEUdotLocalDC
        
        # Create a new custom object to hold our result.
        $userObject = new-object PSObject

        # Add our data to $contactObject as attributes using the add-member commandlet
        $userObject | add-member -membertype NoteProperty -name "SystemID" -Value $SourceUser.SystemID
        $userObject | add-member -membertype NoteProperty -name "Login" -Value $SourceUser.Login
        $userObject | add-member -membertype NoteProperty -name "OldPrimaryEmail" -Value $SourceUser.OldPrimaryEmail
        $userObject | add-member -membertype NoteProperty -name "ElementDN" -Value $SourceUser.ElementDN
        $userObject | add-member -membertype NoteProperty -name "ExovaEUDN" -Value $TargetUser.DistinguishedName
        $userObject | add-member -membertype NoteProperty -name "CurrentUserPrincipalName" -Value $TargetUser.UserPrincipalName
        $userObject | add-member -membertype NoteProperty -name "SamAccountName" -Value $TargetUser.SamAccountName
        $UserObject | add-member -membertype NoteProperty -name "NewEmailAndUsername" -Value $SourceUser.NewEmailAndUsername
        $UserObject | add-member -membertype NoteProperty -name "Notes" -Value $SourceUser.Notes

        # Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
        $resultsarray += $userObject

        # Set-ADUser $SourceUser -Mail $NewMailUPN -UserPrincipalName $NewMailUPN -Server $ExovaEUdotLocalDC

    }
    
    # Export to CSV 
    $resultsarray| Export-csv iCIMS-ExovaOnly-UpdatedEmailAndLoginNames-Round5.csv -notypeinformation


# Use the above data to scan Exova-EU for users who have @element.com already as their mail address (and not that for iCIMS) and so SSO is already broke


    $SourceUsers = Import-CSV iCIMS-ExovaOnly-UpdatedEmailAndLoginNames.csv

    $resultsarray =@()

    foreach ($SourceUser in $SourceUsers) {

        $ExovaUser = $SourceUser.ExovaEUDN
        $TargetUser = Get-ADUser -Property mail -Filter {DistinguishedName -eq $ExovaUser } -Server $ExovaEUdotLocalDC
        
        # Create a new custom object to hold our result.
        $userObject = new-object PSObject

        # Add our data to $contactObject as attributes using the add-member commandlet
        $userObject | add-member -membertype NoteProperty -name "SystemID" -Value $SourceUser.SystemID
        $userObject | add-member -membertype NoteProperty -name "Login" -Value $SourceUser.Login
        $userObject | add-member -membertype NoteProperty -name "OldPrimaryEmail" -Value $SourceUser.OldPrimaryEmail
        $userObject | add-member -membertype NoteProperty -name "ElementDN" -Value $SourceUser.ElementDN
        $userObject | add-member -membertype NoteProperty -name "ExovaEUDN" -Value $SourceUser.ExovaEUDN
        $userObject | add-member -membertype NoteProperty -name "CurrentUserPrincipalName" -Value $SourceUser.UserPrincipalName
        $userObject | add-member -membertype NoteProperty -name "SamAccountName" -Value $SourceUser.SamAccountName
        $UserObject | add-member -membertype NoteProperty -name "NewEmailAndUsername" -Value $SourceUser.NewEmailAndUsername
        $UserObject | add-member -membertype NoteProperty -name "Notes" -Value $SourceUser.Notes
        $UserObject | add-member -membertype NoteProperty -name "Mail" -Value $TargetUser.Mail

        # Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
        $resultsarray += $userObject

        # Set-ADUser $SourceUser -Mail $NewMailUPN -UserPrincipalName $NewMailUPN -Server $ExovaEUdotLocalDC

    }
    
    # Export to CSV 
    $resultsarray| Export-csv iCIMS-ExovaOnly-UpdatedEmailAndLoginNames-MailMatch.csv -notypeinformation









    # Cornerstone Data

    # Get all users without DN and see if they are really there just not by @exova.com address - look for them with @element.com address and add back into spreadsheet

    ## Get DN of user in Exova-EU and build spreadsheet
    $SourceUsers = Import-CSV Cornerstone-ExovaOnly-Source-Round3.csv

    $resultsarray =@()

    Write-Host (Get-Date)

    foreach ($SourceUser in $SourceUsers) {

        $ExovaUser = $SourceUser.Username
          # Needed to use .Username (in the main) and construct UPN and look for possible matches and .OldUserEmail and .NewEmailAndUsername for some users to find them as login name in app is wrong
        $ExovaUser = $ExovaUser.replace("exova.com","exova-eu.local") # replaces exova.com in username to predict the UPN to find the user
        $TargetUser = Get-ADUser -Property mail -Filter {mail -eq $ExovaUser } -Server $ExovaEUdotLocalDC
        # First run, look for "mail" in {} and then later userprincipalname
        
        # Create a new custom object to hold our result.
        $userObject = new-object PSObject
        # Write-Host $SourceUser.Username # to show activity as is now taking some time to run this script

        # Add our data to $contactObject as attributes using the add-member commandlet
        $userObject | add-member -membertype NoteProperty -name "UserID" -Value $SourceUser.UserID
        $userObject | add-member -membertype NoteProperty -name "Username" -Value $SourceUser.Username
        $userObject | add-member -membertype NoteProperty -name "OldUserEmail" -Value $SourceUser.OldUserEmail
        $userObject | add-member -membertype NoteProperty -name "UserFirstName" -Value $SourceUser.UserFirstName
        $userObject | add-member -membertype NoteProperty -name "UserLastName" -Value $SourceUser.UserLastName
        $userObject | add-member -membertype NoteProperty -name "LocalSystemID" -Value $SourceUser.LocalSystemID
        $userObject | add-member -membertype NoteProperty -name "UserLastAccess" -Value $SourceUser.UserLastAccess
        $userObject | add-member -membertype NoteProperty -name "RegionID" -Value $SourceUser.RegionID
        $userObject | add-member -membertype NoteProperty -name "ElementDN" -Value $SourceUser.ElementDN
        $userObject | add-member -membertype NoteProperty -name "ExovaEUDN" -Value $TargetUser.DistinguishedName
        $userObject | add-member -membertype NoteProperty -name "CurrentUserPrincipalName" -Value $TargetUser.UserPrincipalName
        $userObject | add-member -membertype NoteProperty -name "SamAccountName" -Value $TargetUser.SamAccountName
        $UserObject | add-member -membertype NoteProperty -name "NewEmailAndUsername" -Value $SourceUser.NewEmailAndUsername
        $UserObject | add-member -membertype NoteProperty -name "Notes" -Value $SourceUser.Notes

        # Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
        $resultsarray += $userObject

        # Set-ADUser $SourceUser -Mail $NewMailUPN -UserPrincipalName $NewMailUPN -Server $ExovaEUdotLocalDC

    }
    
    # Export to CSV 
    $resultsarray| Export-csv Cornerstone-ExovaOnly-Results-Round8.csv -notypeinformation
    Write-Host (Get-Date)


# Use the above data to scan Exova-EU for users who have @element.com already as their mail address (and not that for Cornerstone) and so SSO is already broke


    $SourceUsers = Import-CSV Cornerstone-ExovaOnly-UpdatedEmailAndLoginNames.csv

    $resultsarray =@()

    foreach ($SourceUser in $SourceUsers) {

        $ExovaUser = $SourceUser.ExovaEUDN
        $TargetUser = Get-ADUser -Property mail -Filter {DistinguishedName -eq $ExovaUser } -Server $ExovaEUdotLocalDC
        
        # Create a new custom object to hold our result.
        $userObject = new-object PSObject

        # Add our data to $contactObject as attributes using the add-member commandlet
        $userObject | add-member -membertype NoteProperty -name "UserID" -Value $SourceUser.UserID
        $userObject | add-member -membertype NoteProperty -name "Username" -Value $SourceUser.Username
        $userObject | add-member -membertype NoteProperty -name "OldUserEmail" -Value $SourceUser.OldUserEmail
        $userObject | add-member -membertype NoteProperty -name "UserFirstName" -Value $SourceUser.UserFirstName
        $userObject | add-member -membertype NoteProperty -name "UserLastName" -Value $SourceUser.UserLastName
        $userObject | add-member -membertype NoteProperty -name "LocalSystemID" -Value $SourceUser.LocalSystemID
        $userObject | add-member -membertype NoteProperty -name "UserLastAccess" -Value $SourceUser.UserLastAccess
        $userObject | add-member -membertype NoteProperty -name "RegionID" -Value $SourceUser.RegionID
        $userObject | add-member -membertype NoteProperty -name "ElementDN" -Value $SourceUser.ElementDN
        $userObject | add-member -membertype NoteProperty -name "ExovaEUDN" -Value $SourceUser.ExovaEUDN
        $userObject | add-member -membertype NoteProperty -name "CurrentUserPrincipalName" -Value $TargetUser.UserPrincipalName
        $userObject | add-member -membertype NoteProperty -name "SamAccountName" -Value $SourceUser.SamAccountName
        $UserObject | add-member -membertype NoteProperty -name "NewEmailAndUsername" -Value $SourceUser.NewEmailAndUsername
        $UserObject | add-member -membertype NoteProperty -name "Notes" -Value $SourceUser.Notes
        $UserObject | add-member -membertype NoteProperty -name "Mail" -Value $TargetUser.Mail

        # Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
        $resultsarray += $userObject

        # Set-ADUser $SourceUser -Mail $NewMailUPN -UserPrincipalName $NewMailUPN -Server $ExovaEUdotLocalDC

    }
    
    # Export to CSV 
    $resultsarray| Export-csv Cornerstone-ExovaOnly-UpdatedEmailAndLoginNames-MailMatch.csv -notypeinformation








    # Filter results by records with ExovaDN and add to the final results spreadsheet and then remove duplicates.
        # Done mail + username Round2
        # Done mail + old Round2
        # Done mail + NewEmailAndUsername Round4
        # Done UPN + username 5 >> this one finds many unique hits
        # Done UPN + old 6 
        # Doing UPN + NewEmailAndUsername 7

    # So will return a spreadsheet without the extra source data.

# Now to update the users from the CSV rather than writing the CSV

# Need to update users (archive first) in Exova-EU with mail and UPN from above CSV

    #iCIMS Spreadsheet
    
    cd C:\ExovaRebrand
    Import-Module ActiveDirectory

    $ExovaEUdotLocalDC = "eudc-ad-03.exova-eu.local"

    $SourceUsers = Import-CSV ExovaOnly-ADDITIONAL1-UpdatedEmailAndLoginNames-iCIMS.csv
    # Needs to be run in Exova-EU domain directly
    foreach ($SourceUser in $SourceUsers) {

       # Write-Host $SourceUser.ExovaEUDN
        Get-ADUser $SourceUser.ExovaEUDN -Server $ExovaEUdotLocalDC -Property Mail | Select Name,Mail,UserPrincipalName # Cannot append to CSV as version of PS is too old on these machines
       # Set-ADUser $SourceUser.ExovaEUDN -EmailAddress $SourceUser.NewEmailAndUsername -UserPrincipalName $SourceUser.NewEmailAndUsername -Server $ExovaEUdotLocalDC 

    }
    
    
        #Cornerstone Spreadsheet
    
    cd C:\ExovaRebrand
    Import-Module ActiveDirectory

    $ExovaEUdotLocalDC = "eudc-ad-03.exova-eu.local"

    $SourceUsers = Import-CSV ExovaOnly-ADDITIONAL1-UpdatedEmailAndLoginNames-CSOD.csv
    # Needs to be run in Exova-EU domain directly
    foreach ($SourceUser in $SourceUsers) {

       # Write-Host $SourceUser.ExovaEUDN
        Get-ADUser $SourceUser.ExovaEUDN -Server $ExovaEUdotLocalDC -Property Mail | Select Name,Mail,UserPrincipalName # Cannot append to CSV as version of PS is too old on these machines
        # $exoc

    }




# Report on all users in Exova-EU that still have @exova.com mail addresses
    Set-ADServerSettings -ViewEntireForest $true
    cd C:\ExovaRebrand
    $ElementEUDC = "usnjdcs040.eu.sys.element.com"
    $ExovaEUDC = "eudc-gad-01.exova.com"
    $ExovaEUdotLocalDC = "eudc-ad-03.exova-eu.local"

$SourceUsers = Get-ADUser -Property mail -Filter {mail -like "*@exova.com" }
$resultsarray =@()

foreach ($SourceUser in $SourceUsers) {

    $userObject = new-object PSObject
    
    $userObject | add-member -membertype NoteProperty -name "UserPrincipalName" -Value $SourceUser.UserPrincipalName
    $userObject | add-member -membertype NoteProperty -name "SamAccountName" -Value $SourceUser.SamAccountName
    $userObject | add-member -membertype NoteProperty -name "DistinguishedName" -Value $SourceUser.DistinguishedName
    $userObject | add-member -membertype NoteProperty -name "Mail" -Value $SourceUser.Mail

    $resultsarray += $userObject

    }
    
    # Export to CSV 
    $resultsarray| Export-csv Exova-EU-Users-Still-ExovaDotCom.csv -notypeinformation



    # See if above have a mailbox
    cd C:\ExovaRebrand
    $SourceUsers = Import-CSV Exova-EU-Users-Still-ExovaDotCom.csv
    $resultsarray =@()
   foreach ($SourceUser in $SourceUsers) {

        $userObject = new-object PSObject

        $targetUser = Get-Recipient $SourceUser.mail
    
        $userObject | add-member -membertype NoteProperty -name "ExovaEUDN" -Value $SourceUser.DistinguishedName
        $userObject | add-member -membertype NoteProperty -name "ExovaEUSamAccountName" -Value $SourceUser.SamAccountName
        $userObject | add-member -membertype NoteProperty -name "Name" -Value $targetUser.name
        $userObject | add-member -membertype NoteProperty -name "RecipientType" -Value $targetUser.recipienttype
        $userObject | add-member -membertype NoteProperty -name "Mail" -Value $SourceUser.mail
        $userObject | add-member -membertype NoteProperty -name "PrimarySmtpAddress" -Value $targetUser.PrimarySmtpAddress
        $userObject | add-member -membertype NoteProperty -name "OrganizationalUnit" -Value $targetUser.OrganizationalUnit

        $resultsarray += $userObject

    }


    $resultsarray| Export-csv Exova-EU-Users-Still-ExovaDotCom-WithMailObject.csv -notypeinformation


# Then LINK them and then SYNC them. Must do it this way around
# "Non Domain User" - see email question for if this is just for exova.com mailbox

    

    
cd C:\ExovaRebrand
$remoteEUdotLocal = Get-Credential -Message "Exova-EU Migration Account (Exova-EU.local)" -UserName "exova-eu\Brian.Reid"
$ExovaEUdotLocalDC = "eudc-ad-03.exova-eu.local"
$ElementEUDC = "usnjdcs040.eu.sys.element.com"
$ElementUSDC = "usnjdcs030.us.sys.element.com"
$ElementMEDC = "usnjdcs050.me.sys.element.com"
$ElementAPDC = "usnjdcs060.ap.sys.element.com"
$ElementCNDC = "usnjdcs070.cn.sys.element.com"
$ExovaEUDC = "eudc-gad-01.exova.com"
$ExovaNADC = "dc-ad02.bcamericas.com"

# Done
# Move all disabled accounts 
# Search-ADAccount -UsersOnly -AccountDisabled -SearchBase "ou=enabled,ou=staging,ou=accounts,dc=eu,dc=sys,dc=element,dc=com" -Server $ElementEUDC | Move-ADObject -TargetPath "ou=staging,ou=accounts,dc=eu,dc=sys,dc=element,dc=com" -Server $ElementEUDC

# Move Exova-EU and Exova.com user accounts into the Sync OU if they are archived users (first test)
# Move above if they are Warrington or Tolouse (as sync OUs exist for them)
# Move others who are disabled - do I need to make OUs for them?


# Link mailboxes 
# Updating linked account disables the source mailbox and sets the source UPN to come from exova-eu
# Therefore care when doing intentionally enabled accounts (re-enable at end)
# Link users in small batches and sync them to cloud. If disabled in Element can do in bigger batches. If enabled ensure remain enabled and treat with care
# Dont need to do all users when iCIMS/Cornerstone is fixed - only cannot start until that time.
    # Link | Enable if needed | Move object into sync in exova and exova-eu and sync then and resolve errors and repeat (2 days work)

    ##### Takes about 20 seconds per user when working. If errors its becuase un matching samAccountName - user should be found eventually - script tries 3 different ways
    #####
    ##### PROBLEMS
    ##### Metech users should not be run here as they are not to be linked to these accounts. 
    ##### Dont sync them either

# This needs testing and runs on Element Exchange Server (so need to copy CSV to this location as well)
#$Users = Import-CSV iCIMS-Exova-Final.csv # iCIMS Sourced
$Users = Import-CSV CSOD-Exova-Final.csv  # CSOD Sourced

Set-ADServerSettings -ViewEntireForest $true
ForEach ($User in $Users) {
    $ElementUser = Get-User $User.ElementDN
    #$LinkedAccount = "EXOVA-EU\" + $User.ExovaEUSamAccountName # iCIMS sourced
    $LinkedAccount = "EXOVA-EU\" + $User.SamAccountName # CSOD Sourced

    Write-Host $LinkedAccount

    if ($ElementUser.DistinguishedName -ilike "*DC=ap,DC=sys,DC=element,DC=com*") {
        Set-User $ElementUser -LinkedMasterAccount $LinkedAccount -LinkedDomainController $ExovaEUdotLocalDC -LinkedCredential $remoteEUdotLocal -DomainController $ElementAPDC
        }
    if ($ElementUser.DistinguishedName -ilike "*DC=me,DC=sys,DC=element,DC=com*") {
        Set-User $ElementUser -LinkedMasterAccount $LinkedAccount -LinkedDomainController $ExovaEUdotLocalDC -LinkedCredential $remoteEUdotLocal -DomainController $ElementMEDC
        }
    if ($ElementUser.DistinguishedName -ilike "*DC=eu,DC=sys,DC=element,DC=com*") {
        Set-User $ElementUser -LinkedMasterAccount $LinkedAccount -LinkedDomainController $ExovaEUdotLocalDC -LinkedCredential $remoteEUdotLocal -DomainController $ElementEUDC
        }
    if ($ElementUser.DistinguishedName -ilike "*DC=us,DC=sys,DC=element,DC=com*") {
        Set-User $ElementUser -LinkedMasterAccount $LinkedAccount -LinkedDomainController $ExovaEUdotLocalDC -LinkedCredential $remoteEUdotLocal -DomainController $ElementUSDC
        }
    if ($ElementUser.DistinguishedName -ilike "*DC=cn,DC=sys,DC=element,DC=com*") {
        Set-User $ElementUser -LinkedMasterAccount $LinkedAccount -LinkedDomainController $ExovaEUdotLocalDC -LinkedCredential $remoteEUdotLocal -DomainController $ElementCNDC
        }



    # Enable the account (if it has a valid password this will work, and it was probably enabled before). If it was previously disabled it will error and stay disabled
    # Can take over a minute per user to do this!

    if ($ElementUser.DistinguishedName -ilike "*DC=ap,DC=sys,DC=element,DC=com*") {
        Enable-ADAccount $ElementUser.DistinguishedName -Server $ElementAPDC
        }
    if ($ElementUser.DistinguishedName -ilike "*DC=me,DC=sys,DC=element,DC=com*") {
        Enable-ADAccount $ElementUser.DistinguishedName -Server $ElementMEDC
        }
    if ($ElementUser.DistinguishedName -ilike "*DC=eu,DC=sys,DC=element,DC=com*") {
        Enable-ADAccount $ElementUser.DistinguishedName -Server $ElementEUDC
        }
    if ($ElementUser.DistinguishedName -ilike "*DC=us,DC=sys,DC=element,DC=com*") {
        Enable-ADAccount $ElementUser.DistinguishedName -Server $ElementUSDC
        }
    if ($ElementUser.DistinguishedName -ilike "*DC=cn,DC=sys,DC=element,DC=com*") {
        Enable-ADAccount $ElementUser.DistinguishedName -Server $ElementCNDC
        }

    }

    # Done