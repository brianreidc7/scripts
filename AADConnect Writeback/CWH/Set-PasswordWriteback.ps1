$passwordOU = "DC=cwh,DC=pri" #[you can scope this down to a specific OU]
$accountName = "hsgnt\MSOL_66aee734aedf" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is often an account in the form of MSOL_number or AAD_number].
 
$cmd = "dsacls.exe '$passwordOU' /I:S /G '`"$accountName`":CA;`"Reset Password`";user'"
Invoke-Expression $cmd | Out-Null
 
$cmd = "dsacls.exe '$passwordOU' /I:S /G '`"$accountName`":CA;`"Change Password`";user'"
Invoke-Expression $cmd | Out-Null
 
$cmd = "dsacls.exe '$passwordOU' /I:S /G '`"$accountName`":WP;lockoutTime;user'"
Invoke-Expression $cmd | Out-Null
 
$cmd = "dsacls.exe '$passwordOU' /I:S /G '`"$accountName`":WP;pwdLastSet;user'"
Invoke-Expression $cmd | Out-Null