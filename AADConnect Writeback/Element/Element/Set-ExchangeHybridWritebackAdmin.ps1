$accountName = "sys\aadsync" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is often an account in the form of MSOL_number or AAD_number].
$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=sys,DC=element,DC=com"
 
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


$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=ent,DC=sys,DC=element,DC=com"
 
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


$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=us,DC=sys,DC=element,DC=com"
 
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


$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=eu,DC=sys,DC=element,DC=com"
 
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


$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=ap,DC=sys,DC=element,DC=com"
 
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


$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=me,DC=sys,DC=element,DC=com"
 
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


$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=ven,DC=sys,DC=element,DC=com"
 
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


$AdminSDHolder = "CN=AdminSDHolder,CN=System,DC=cn,DC=sys,DC=element,DC=com"
 
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


