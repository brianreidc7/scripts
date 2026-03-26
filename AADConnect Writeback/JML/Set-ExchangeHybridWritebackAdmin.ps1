$accountName = "jml\MSOL_16d1fff0150f" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is an account usually in the form of AAD_number or MSOL_number].
$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=jml,DC=local"
 
$cmd = "dsacls '$AdminSDHolder' /G '`"$accountName`":WP;proxyAddresses'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$AdminSDHolder' /G '`"$accountName`":WP;msExchUCVoiceMailSettings'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$AdminSDHolder' /G '`"$accountName`":WP;msExchUserHoldPolicies'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$AdminSDHolder' /G '`"$accountName`":WP;msExchArchiveStatus'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$AdminSDHolder' /G '`"$accountName`":WP;msExchSafeSendersHash'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$AdminSDHolder' /G '`"$accountName`":WP;msExchBlockedSendersHash'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$AdminSDHolder' /G '`"$accountName`":WP;msExchSafeRecipientsHash'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$AdminSDHolder' /G '`"$accountName`":WP;msDS-ExternalDirectoryObjectID'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$AdminSDHolder' /G '`"$accountName`":WP;publicDelegates'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$AdminSDHolder' /G '`"$accountName`":WP;msExchDelegateLinkList'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$AdminSDHolder' /G '`"$accountName`":WP;msExchUserHoldPolicies'"
Invoke-Expression $cmd | Out-Null

