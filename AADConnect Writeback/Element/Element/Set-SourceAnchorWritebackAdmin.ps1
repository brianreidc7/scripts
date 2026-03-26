$accountName = "sys\aadsync" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is often an account in the form of MSOL_number or AAD_number].

$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;ms-ds-consistencyGuid'"
Invoke-Expression $cmd | Out-Null

$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=ent,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;ms-ds-consistencyGuid'"
Invoke-Expression $cmd | Out-Null

$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=eu,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;ms-ds-consistencyGuid'"
Invoke-Expression $cmd | Out-Null

$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=us,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;ms-ds-consistencyGuid'"
Invoke-Expression $cmd | Out-Null

$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=ven,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;ms-ds-consistencyGuid'"
Invoke-Expression $cmd | Out-Null

$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=me,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;ms-ds-consistencyGuid'"
Invoke-Expression $cmd | Out-Null

$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=ap,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;ms-ds-consistencyGuid'"
Invoke-Expression $cmd | Out-Null

$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=cn,DC=sys,DC=element,DC=com"
 
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;ms-ds-consistencyGuid'"
Invoke-Expression $cmd | Out-Null

