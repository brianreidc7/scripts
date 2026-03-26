$accountName = "jml\MSOL_16d1fff0150f" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is an account usually in the form of AAD_number or MSOL_number].
$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=jml,DC=local"
 
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":CA;`"Reset Password`"'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":CA;`"Change Password`"'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;lockoutTime'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;pwdLastSet'"
Invoke-Expression $cmd | Out-Null