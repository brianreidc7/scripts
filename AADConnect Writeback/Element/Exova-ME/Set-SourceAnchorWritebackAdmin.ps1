$accountName = "exova-me\MSOL_ff3f1e7e5b34" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is an account usually in the form of AAD_number or MSOL_number].
$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=exova-me,DC=local"
 
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;ms-ds-consistencyGuid'"
Invoke-Expression $cmd | Out-Null