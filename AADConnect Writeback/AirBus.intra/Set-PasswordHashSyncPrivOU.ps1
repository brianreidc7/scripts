$accountName = "airbus\MSOL_098a750bbf2d" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is often an account in the form of MSOL_number or AAD_number].
$DomainDN = "OU=Privileged Access Control,OU=Air Business Limited,DC=airbus,DC=intra"
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes`";'"
Invoke-Expression $cmd | Out-Null
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes All`";'"
Invoke-Expression $cmd | Out-Null