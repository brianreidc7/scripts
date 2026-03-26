### Process disabled staging users



### Get User Properties for each OU

get-user -OrganizationalUnit "us.sys.element.com/Accounts/Staging" | select Name,UserPrincipalName,IsLinked,LinkedMasterAccount,FirstName,LastName,WindowsEmailAddress,RecipientType | Export-CSV "US Staging OU Only (Bulk Sync).csv" -NoTypeInformation
get-user -OrganizationalUnit "me.sys.element.com/Accounts/Staging" | select Name,UserPrincipalName,IsLinked,LinkedMasterAccount,FirstName,LastName,WindowsEmailAddress,RecipientType | Export-CSV "ME Staging OU Only (Bulk Sync).csv" -NoTypeInformation
get-user -OrganizationalUnit "ap.sys.element.com/Accounts/Staging" | select Name,UserPrincipalName,IsLinked,LinkedMasterAccount,FirstName,LastName,WindowsEmailAddress,RecipientType | Export-CSV "AP Staging OU Only (Bulk Sync).csv" -NoTypeInformation
get-user -OrganizationalUnit "cn.sys.element.com/Accounts/Staging" | select Name,UserPrincipalName,IsLinked,LinkedMasterAccount,FirstName,LastName,WindowsEmailAddress,RecipientType | Export-CSV "CN Staging OU Only (Bulk Sync).csv" -NoTypeInformation




### Get disabled or not. Input it the above CSV

$SourceCSV = "AP Staging OU Only (Bulk Sync).csv"
$IsEnabled = Import-CSV $SourceCSV
$resultsarray =@()

ForEach ($SourceUser in $IsEnabled) {

        $userObject = new-object PSObject
        $TargetUser = $SourceUser.UserPrincipalName
        $ExovaEUTargetUser = Get-ADUser -Filter {UserPrincipalName -eq $TargetUser}

        # Add our data to $contactObject as attributes using the add-member commandlet
        $userObject | add-member -membertype NoteProperty -name "Name" -Value $SourceUser.Name
        $userObject | add-member -membertype NoteProperty -name "UserPrincipalName" -Value $SourceUser.UserPrincipalName
        $userObject | add-member -membertype NoteProperty -name "IsLinked" -Value $SourceUser.IsLinked
        $userObject | add-member -membertype NoteProperty -name "LinkedMasterAccount" -Value $SourceUser.LinkedMasterAccount
        $userObject | add-member -membertype NoteProperty -name "FirstName" -Value $SourceUser.FirstName
        $userObject | add-member -membertype NoteProperty -name "LastName" -Value $SourceUser.LastName
        $userObject | add-member -membertype NoteProperty -name "WindowsEmailAddress" -Value $SourceUser.WindowsEmailAddress
        $userObject | add-member -membertype NoteProperty -name "RecipientType" -Value $SourceUser.RecipientType
        $userObject | add-member -membertype NoteProperty -name "Enabled" -Value $SourceUser.Enabled
        $userObject | add-member -membertype NoteProperty -name "ExovaEUDN" -Value $ExovaEUTargetUser.Distinguishedname
        $userObject | add-member -membertype NoteProperty -name "ExovaEUSamAccountName" -Value $ExovaEUTargetUser.SamAccountName
        $userObject | add-member -membertype NoteProperty -name "ExovaEUUserPrincipalName" -Value $ExovaEUTargetUser.UserPrincipalName

        # Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
        $resultsarray += $userObject
	
}

 $resultsarray| Export-csv $SourceCSV -notypeinformation




 ### Filter results.
 # Disabled only / MailUser only / Unlinked = to link them 
 # & wrong UPN to fix
 # All done in Excel cmdlets stitching


# Started this - AP file is ready, but cancelled at soon after starting
# Cmdlet for linking only. Not UPN