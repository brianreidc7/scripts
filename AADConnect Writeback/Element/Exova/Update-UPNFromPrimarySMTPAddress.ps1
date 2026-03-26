# Get all the users who have proxyAddresses under the eu.sys.element.com domain
# This updates all users without checking - see below for script that updates based on a CSV (exported those users with licences for example)
foreach ($user in (Get-ADUser -SearchBase "OU=Staging,OU=Accounts,DC=eu,DC=sys,DC=element,DC=com" -LdapFilter '(proxyAddresses=*)')) {
	# Grab the primary SMTP address
	$address = Get-ADUser $user -Properties proxyAddresses | Select -Expand proxyAddresses | Where {$_ -clike "SMTP:*"}
	# Remove the protocol specification from the start of the address
	$newUPN = $address.SubString(5)
	# Update the user with their new UPN
	Set-ADUser $user -UserPrincipalName $newUPN
    Write-Host $user "===" $newUPN
}


#Write UPN to the screen (user, current UPN and proposed newUPN based on proxyAddress
$resultsarray =@()
foreach ($user in (Get-ADUser -SearchBase "OU=exova-staging-cloud,OU=Accounts,DC=eu,DC=sys,DC=element,DC=com" -LdapFilter '(proxyAddresses=*)')) {
	# Grab the primary SMTP address
	$address = Get-ADUser $user -Properties proxyAddresses | Select -Expand proxyAddresses | Where {$_ -clike "SMTP:*"}
    $currentUPN = Get-ADUser $user -Properties UserPrincipalName
	# Remove the protocol specification from the start of the address
	$newUPN = $address.SubString(5)
	# Update the user with their new UPN

    # Create a new custom object to hold our result.
    $userObject = new-object PSObject

    # Add our data to $contactObject as attributes using the add-member commandlet
    $userObject | add-member -membertype NoteProperty -name "Name" -Value $user.name
    $userObject | add-member -membertype NoteProperty -name "SamAccountName" -Value $user.samAccountName
    $userObject | add-member -membertype NoteProperty -name "UserPrincipalName" -Value $currentUPN.UserPrincipalName
    $userObject | add-member -membertype NoteProperty -name "NewUPN" -Value $newUPN

    # Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
    $resultsarray += $userObject
}
$resultsarray| Export-csv ElementUPNs.csv -notypeinformation


# Update all users in CSV to new UPN (from above script) that have licence info added - update users with no cloud licence
# CSV should only contain unlicenced users
$users = Import-Csv -Path C:\temp\element-exova-cloud-upn-check-nolicence.csv
foreach ($user in $users) {
	# Update the user with their new UPN
	Set-ADUser $user.SamAccountName -UserPrincipalName $user.newUPN
    Write-Host $user.SamAccountName "===" $user.newUPN
}
