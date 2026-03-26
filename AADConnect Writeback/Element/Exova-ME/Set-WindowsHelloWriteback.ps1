$accountName = "exova-me\MSOL_ff3f1e7e5b34" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is often an account in the form of MSOL_number or AAD_number].
$ForestDN = "DC=exova-me,DC=local"
 
$cmd = "dsacls '$ForestDN' /I:S /G '`"$accountName`":WP;msDs-KeyCredentialslLink;user'"
Invoke-Expression $cmd | Out-Null

# This script requires the Active Directory 2016 Schema installed (a 2016 Domain Controller install will do this, but you can update the schema without adding a new DC)