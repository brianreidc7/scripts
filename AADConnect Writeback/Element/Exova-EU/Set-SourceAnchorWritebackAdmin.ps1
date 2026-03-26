$accountName = "exova-eu\aadsync" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is an account usually in the form of AAD_number or MSOL_number].
$AdminSDHolder = "CN=AdminSDHolder,CN=System,dc=exova-eu,dc=local"
 
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;ms-ds-consistencyGuid'"
Invoke-Expression $cmd | Out-Null
