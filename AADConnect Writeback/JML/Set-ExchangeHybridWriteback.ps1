$accountName = "jml\MSOL_16d1fff0150f" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is an account usually in the form of AAD_number or MSOL_number].
$HybridOU = "DC=jml,DC=local"
 
#Object type: user
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;proxyAddresses;user'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchUCVoiceMailSettings;user'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchUserHoldPolicies;user'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchArchiveStatus;user'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchSafeSendersHash;user'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchBlockedSendersHash;user'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchSafeRecipientsHash;user'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msDS-ExternalDirectoryObjectID;user'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;publicDelegates;user'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchDelegateLinkList;user'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchUserHoldPolicies;user'"
Invoke-Expression $cmd | Out-Null

 
#Object type: iNetOrgPerson
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;proxyAddresses;iNetOrgPerson'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchUCVoiceMailSettings;iNetOrgPerson'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchUserHoldPolicies;iNetOrgPerson'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchArchiveStatus;iNetOrgPerson'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchSafeSendersHash;iNetOrgPerson'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchBlockedSendersHash;iNetOrgPerson'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchSafeRecipientsHash;iNetOrgPerson'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msDS-ExternalDirectoryObjectID;iNetOrgPerson'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;publicDelegates;iNetOrgPerson'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchDelegateLinkList;iNetOrgPerson'"
Invoke-Expression $cmd | Out-Null
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;msExchUserHoldPolicies;iNetOrgPerson'"
Invoke-Expression $cmd | Out-Null

 
#Object type: group
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;proxyAddresses;group'"
Invoke-Expression $cmd | Out-Null
 
#Object type: contact
$cmd = "dsacls '$HybridOU' /I:S /G '`"$accountName`":WP;proxyAddresses;contact'"
Invoke-Expression $cmd | Out-Null