$accountName = "sys\aadsync" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is often an account in the form of MSOL_number or AAD_number].
$DomainDN = "DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes`";'"
Invoke-Expression $cmd | Out-Null
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes All`";'"
Invoke-Expression $cmd | Out-Null

$DomainDN = "DC=cn,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes`";'"
Invoke-Expression $cmd | Out-Null
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes All`";'"
Invoke-Expression $cmd | Out-Null

$DomainDN = "DC=me,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes`";'"
Invoke-Expression $cmd | Out-Null
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes All`";'"
Invoke-Expression $cmd | Out-Null

$DomainDN = "DC=ap,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes`";'"
Invoke-Expression $cmd | Out-Null
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes All`";'"
Invoke-Expression $cmd | Out-Null

$DomainDN = "DC=ven,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes`";'"
Invoke-Expression $cmd | Out-Null
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes All`";'"
Invoke-Expression $cmd | Out-Null

$DomainDN = "DC=us,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes`";'"
Invoke-Expression $cmd | Out-Null
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes All`";'"
Invoke-Expression $cmd | Out-Null

$DomainDN = "DC=eu,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes`";'"
Invoke-Expression $cmd | Out-Null
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes All`";'"
Invoke-Expression $cmd | Out-Null

$DomainDN = "DC=ent,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes`";'"
Invoke-Expression $cmd | Out-Null
 
$cmd = "dsacls.exe '$DomainDN' /G '`"$accountName`":CA;`"Replicating Directory Changes All`";'"
Invoke-Expression $cmd | Out-Null

