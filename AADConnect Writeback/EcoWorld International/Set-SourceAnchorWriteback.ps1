$accountName = "residential\sa_aadconnect" #[this is the account that will be used by Azure AD Connect Sync to manage objects in the directory, this is often an account in the form of MSOL_number or AAD_number].
$ForestDN = "DC=willmottresidential,DC=co,DC=uk"
 
$cmd = "dsacls '$ForestDN' /I:S /G '`"$accountName`":WP;ms-ds-consistencyGuid;user'"
Invoke-Expression $cmd | Out-Null
