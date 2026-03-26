$accountName = "domain\aad_account"
$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=contoso,DC=com"
 
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


