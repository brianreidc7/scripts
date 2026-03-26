<#
	.SYNOPSIS
	Convert-MailboxToRemoteMailbox will convert an existing mailbox to a RemoteMailbox. Based on work by Nero Blanco at https://www.neroblanco.co.uk/2015/06/convert-mailbox-to-mailuser/

	.DESCRIPTION
	Convert-MailboxToRemoteMailbox will convert an existing mailbox to a RemoteMailbox.  It will copy the main Exchange properties from the mailbox to the RemoteMailbox

	.PARAMETER Identity
	The Identity parameter specifies the identity of the mailbox. You can use one of the following values:
	* GUID
	* Distinguished name (DN)
	* Display name
	* Domain\Account
	* User principal name (UPN)
	* LegacyExchangeDN
	* SmtpAddress
	* Alias

	.PARAMETER DomainController
        The DomainController parameter specifies the fully qualified domain name (FQDN) of the domain controller that retrieves data from Active Directory.

	.Parameter ExternalEmailAddress <ProxyAddress>
	The ExternalEmailAddress parameter specifies an e-mail address outside the organization. E-mail messages sent to the user are sent to this external address.

	.EXAMPLE
	This example converts the mailbox Chris to a RemoteMailbox with the external email address Chris@contoso.com

	.\Convert-MailboxToRemoteMailbox.ps1 -Identity Chris -ExternalEmailAddress Chris@contoso.com

	.EXAMPLE
	This example converts the mailbox Chris to a RemoteMailbox with the external email address Chris@contoso.com using the specified domain controller

	.\Convert-MailboxToRemoteMailbox.ps1 -Identity Chris -ExternalEmailAddress Chris@contoso.com -DomainController MyDC

    .EXAMPLE
    This example processes a CSV file with one column called Identity to process users in bulk

    Import-CSV .\csvfile.csv | foreach-object { .\Convert-MailboxToMailuser.ps1 -identity $_.Identity -DomainController $_.DomainController -ExternalEmailAddress $_.ExternalEmailAddress }
#>

[CmdletBinding(
	SupportsShouldProcess=$true  # removed ConfirmImapact High as this causes confirmation on each run,
	#ConfirmImpact="High"
)]
Param
(
	[Parameter(Mandatory=$true,Position=0)]
	[String]$Identity,
	[Parameter(Mandatory=$false,Position=1)]
	[String]$ExternalEmailAddress,
	[Parameter(Mandatory=$false,Position=2)]
	[String]$DomainController
)
PROCESS
{

	$parameters = @{}
	if( $DomainController ) {
		$parameters.Add( 'DomainController', $DomainController )
		Write-Verbose ( 'Using Domain Controller "{0}"' -f $DomainController )
	}

	Write-Verbose ( 'Trying to find mailbox for "{0}"' -f $Identity )
	$OldMailbox = get-mailbox $Identity @parameters -ErrorAction SilentlyContinue

	if( ($OldMailbox) -and -not ($OldMailbox.count -gt 1) ) {
		if( $pscmdlet.ShouldProcess($OldMailbox.DistinguishedName) ) {
			$parameters.add( 'Identity', $OldMailbox.DistinguishedName )

			Write-Verbose ( 'Disabling mailbox "{0}"' -f $OldMailbox.DistinguishedName )
			Disable-mailbox @parameters -confirm:$false

            #$parameters.add( 'ExternalEmailAddress', $ExternalEmailAddress )
            $ExternalEmailAddress = $OldMailbox.Alias + "@tenant.mail.onmicrosoft.com" 

			$parameters.add( 'RemoteRoutingAddress', $ExternalEmailAddress )

			Write-Verbose ( 'Enabling RemoteMailbox "{0}"' -f $OldMailbox.DistinguishedName )
			Enable-RemoteMailbox @parameters

			$parameters.add( 'EmailAddressPolicyEnabled', $OldMailbox.EmailAddressPolicyEnabled )

			$EmailAddresses= @( 'x500:{0}' -f $OldMailbox.legacyExchangeDN )

			if( $OldMailbox.EmailAddressPolicyEnabled ) { 
				foreach( $EmailAddress in $OldMailbox.EmailAddresses ) {
					if( $EmailAddress -like 'smtp:*' ) {
						$EmailAddresses += $EmailAddress.ToString().ToLower()
					} else {
						$EmailAddresses += $EmailAddress.ToString()
					}
				}

				$parameters.add( 'EmailAddresses', @{Add=$EmailAddresses} )
			} else {
				foreach( $EmailAddress in $OldMailbox.EmailAddresses ) {
					$EmailAddresses += $EmailAddress.ToString()
				}

				$parameters.add( 'EmailAddresses', $EmailAddresses )
			}

			$parameters.add( 'DisplayName', $OldMailbox.DisplayName )
			$parameters.add( 'Alias', $OldMailbox.Alias )
			$parameters.add( 'CustomAttribute1', $OldMailbox.CustomAttribute1 )
			$parameters.add( 'CustomAttribute2', $OldMailbox.CustomAttribute2 )
			$parameters.add( 'CustomAttribute3', $OldMailbox.CustomAttribute3 )
			$parameters.add( 'CustomAttribute4', $OldMailbox.CustomAttribute4 )
			$parameters.add( 'CustomAttribute5', $OldMailbox.CustomAttribute5 )
			$parameters.add( 'CustomAttribute6', $OldMailbox.CustomAttribute6 )
			$parameters.add( 'CustomAttribute7', $OldMailbox.CustomAttribute7 )
			$parameters.add( 'CustomAttribute8', $OldMailbox.CustomAttribute8 )
			$parameters.add( 'CustomAttribute9', $OldMailbox.CustomAttribute9 )
			$parameters.add( 'CustomAttribute10', $OldMailbox.CustomAttribute10 )
			$parameters.add( 'CustomAttribute11', $OldMailbox.CustomAttribute11 )
			$parameters.add( 'CustomAttribute12', $OldMailbox.CustomAttribute12 )
			$parameters.add( 'CustomAttribute13', $OldMailbox.CustomAttribute13 )
			$parameters.add( 'CustomAttribute14', $OldMailbox.CustomAttribute14 )
			$parameters.add( 'CustomAttribute15', $OldMailbox.CustomAttribute15 )
			$parameters.add( 'ExtensionCustomAttribute1', $OldMailbox.ExtensionCustomAttribute1 )
			$parameters.add( 'ExtensionCustomAttribute2', $OldMailbox.ExtensionCustomAttribute2 )
			$parameters.add( 'ExtensionCustomAttribute3', $OldMailbox.ExtensionCustomAttribute3 )
			$parameters.add( 'ExtensionCustomAttribute4', $OldMailbox.ExtensionCustomAttribute4 )
			$parameters.add( 'ExtensionCustomAttribute5', $OldMailbox.ExtensionCustomAttribute5 )

            # New attributes to make a remote mailbox
            #msExchPreviousRecipientTypeDetails = 1
            #msExchRecipientDisplayType = -2147483642
            #msExchRecipientSoftDeletedStatus = 0
            #msExchRecipientTypeDetails = 2147483648
            #msExchRemoteRecipientType = 1

			Write-Verbose ( 'Updating RemoteMailbox "{0}"' -f $OldMailbox.DistinguishedName )
			Set-RemoteMailbox @parameters


		}
	} elseif( $OldMailbox.Count -gt 1 ) {
		Write-Host ( 'Multiple mailboxes found for "{0}"' -f $Identity ) -ForegroundColor Red
	} else {
		Write-Host ( 'Unable to find mailbox for "{0}"' -f $Identity ) -ForegroundColor Red
	}
}