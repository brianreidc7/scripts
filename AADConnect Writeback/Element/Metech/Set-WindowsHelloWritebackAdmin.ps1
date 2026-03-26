$accountName = "metech\MSOL_ff3f1e7e5b34" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is an account usually in the form of AAD_number or MSOL_number].
$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=metech,DC=local"
 
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;msDs-KeyCredentialslLink'"
Invoke-Expression $cmd | Out-Null

# This script requires the Active Directory 2016 Schema installed (a 2016 Domain Controller install will do this, but you can update the schema without adding a new DC)

