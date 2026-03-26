$accountName = "residential\sa_aadconnect" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is often an account in the form of MSOL_number or AAD_number].
$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=willmottresidential,DC=co,DC=uk"
 
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":CA;`"Reset Password`"'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":CA;`"Change Password`"'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;lockoutTime'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls.exe '$AdminSDHolder' /G '`"$accountName`":WP;pwdLastSet'"
Invoke-Expression $cmd | Out-Null
