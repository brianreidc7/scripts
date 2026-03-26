$accountName = "jml\MSOL_16d1fff0150f" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is an account usually in the form of AAD_number or MSOL_number].
$DomainDN = "DC=jml,DC=local"
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes`";'"
Invoke-Expression $cmd | Out-Null
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes All`";'"
Invoke-Expression $cmd | Out-Null