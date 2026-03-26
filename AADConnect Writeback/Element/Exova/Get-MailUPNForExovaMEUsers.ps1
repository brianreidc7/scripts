cd C:\ExovaRebrand
#$remoteEUdotLocal = Get-Credential -Message "Exova-EU Migration Account (Exova-EU.local)" -UserName "exova-eu\Brian.Reid"
$ExovaEUdotLocalDC = "eudc-ad-03.exova-eu.local"
$ElementEUDC = "usnjdcs040.eu.sys.element.com"
$ElementUSDC = "usnjdcs030.us.sys.element.com"
$ElementMEDC = "usnjdcs050.me.sys.element.com"
$ElementAPDC = "usnjdcs060.ap.sys.element.com"
$ElementCNDC = "usnjdcs070.cn.sys.element.com"
$ExovaEUDC = "eudc-gad-01.exova.com"
$ExovaNADC = "dc-ad02.bcamericas.com"
$MetechDC = "sedcad04.metech.local"
$ExovaMEDC = "USNJDCS095.exova-me.local"

Import-Module ActiveDirectory

    # iCIMS Data

    ## Get DN of user in Exova-EU and build spreadsheet
    $SourceUsers = Import-CSV iCIMS-ExovaME.csv

    $resultsarray =@()

    foreach ($SourceUser in $SourceUsers) {

        $MetechUser = $SourceUser.Login  
        $TargetUser = Get-ADUser -Properties mail -Filter {mail -eq $MetechUser } -Server $ExovaMEDC
        
        # Create a new custom object to hold our result.
        $userObject = new-object PSObject

        # Add our data to $contactObject as attributes using the add-member commandlet
        $userObject | add-member -membertype NoteProperty -name "SystemID" -Value $SourceUser.SystemID
        $userObject | add-member -membertype NoteProperty -name "Login" -Value $SourceUser.Login
        $userObject | add-member -membertype NoteProperty -name "OldPrimaryEmail" -Value $SourceUser.OldPrimaryEmail
        $userObject | add-member -membertype NoteProperty -name "ElementDN" -Value $SourceUser.ElementDN
        $userObject | add-member -membertype NoteProperty -name "ExovaEUDN" -Value $SourceUser.ExovaEUDN
        $userObject | add-member -membertype NoteProperty -name "CurrentExovaEUUserPrincipalName" -Value $SourceUser.CurrentExovaEUUserPrincipalName
        $userObject | add-member -membertype NoteProperty -name "ExovaEUSamAccountName" -Value $SourceUser.ExovaEUSamAccountName
        $UserObject | add-member -membertype NoteProperty -name "NewEmailAndUsername" -Value $SourceUser.NewEmailAndUsername
        $UserObject | add-member -membertype NoteProperty -name "Notes" -Value $SourceUser.Notes
        $UserObject | add-member -membertype NoteProperty -name "MetechDN" -Value $TargetUser.DistinguishedName
        $UserObject | add-member -membertype NoteProperty -name "MetechOldMail" -Value $TargetUser.Mail

        # Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
        $resultsarray += $userObject


    }
    
    # Export to CSV 
    $resultsarray| Export-csv iCIMS-ExovaMe-Updated.csv -notypeinformation








    # Cornerstone Data

    # Get all users without DN and see if they are really there just not by @exova.com address - look for them with @element.com address and add back into spreadsheet

    ## Get DN of user in Exova-EU and build spreadsheet
    $SourceUsers = Import-CSV CSOD-ExovaME.csv

    $resultsarray =@()

 
    foreach ($SourceUser in $SourceUsers) {

        $ExovaUser = $SourceUser.Username
        $TargetUser = Get-ADUser -Properties mail -Filter {mail -eq $ExovaUser } -Server $ExovaMEDC
        
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
        $userObject | add-member -membertype NoteProperty -name "CurrentUserPrincipalName" -Value $SourceUser.CurrentUserPrincipalName
        $userObject | add-member -membertype NoteProperty -name "SamAccountName" -Value $SourceUser.SamAccountName
        $UserObject | add-member -membertype NoteProperty -name "NewEmailAndUsername" -Value $SourceUser.NewEmailAndUsername
        $UserObject | add-member -membertype NoteProperty -name "Notes" -Value $SourceUser.Notes
        $UserObject | add-member -membertype NoteProperty -name "MetechDN" -Value $TargetUser.DistinguishedName
        $UserObject | add-member -membertype NoteProperty -name "Mail" -Value $TargetUser.Mail

        # Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
        $resultsarray += $userObject


    }
    
    # Export to CSV 
    $resultsarray| Export-csv CSOD-ExovaME-Updated.csv -notypeinformation



    #Update all users (two spreadsheets)

    cd C:\ExovaRebrand
    Import-Module ActiveDirectory
    $ExovaMEDC = "USNJDCS095.exova-me.local"

    $SourceUsers = Import-CSV iCIMS-ExovaME-Updated.csv
 
    foreach ($SourceUser in $SourceUsers) {

       # Write-Host $SourceUser.ExovaEUDN
        Get-ADUser $SourceUser.MetechDN -Server $ExovaMEDC -Property Mail | Select Name,Mail # Cannot append to CSV as version of PS is too old on these machines
        #Set-ADUser $SourceUser.MetechDN -EmailAddress $SourceUser.NewEmailAndUsername -Server $ExovaMEDC 

    }
    
    
        #Cornerstone Spreadsheet
    


    $SourceUsers = Import-CSV CSOD-ExovaME-Updated.csv
    
    foreach ($SourceUser in $SourceUsers) {

       # Write-Host $SourceUser.ExovaEUDN
        Get-ADUser $SourceUser.MetechDN -Server $ExovaMEDC -Property Mail | Select Name,Mail # Cannot append to CSV as version of PS is too old on these machines
        # Set-ADUser $SourceUser.MetechDN -EmailAddress $SourceUser.NewEmailAndUsername -Server $ExovaMEDC 

    }

