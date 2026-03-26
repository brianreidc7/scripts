$DisplayName = "Justin Grier"
$EmailAddress = "justin.grier@pctest.com"
$Message = "Invitation so you can access the 'Element IT Department Team Site' in Teams"
New-AzureADMSInvitation -InvitedUserDisplayName $DisplayName -InvitedUserEmailAddress $EmailAddress -SendInvitationMessage $false -InvitedUserType Member -InviteRedirectUrl https://myapps.microsoft.com
