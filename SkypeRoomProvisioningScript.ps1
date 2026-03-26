#requires -version 3

<#
//    Copyright (c) Microsoft Corporation. All rights reserved.
//    This code is licensed under the Microsoft Public License.
//    THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
//    ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
//    IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
//    PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
//
//    Skype Room Devices Provisioning Script Version 0.33

#>
###############################################################################
#
#    Prerequisites for Office 365
#
#    Connect to Office 365 using Windows PowerShell: http://go.microsoft.com/fwlink/p/?LinkID=614839 
#
#    Connect to Exchange Online using Windows PowerShell: http://go.microsoft.com/fwlink/p/?LinkId=396554
#
#    Connect to Skype for Business Online using Windows PowerShell: http://go.microsoft.com/fwlink/p/?LinkId=691607
#1
###############################################################################

#
# Global Constants
#

$script:strDevice = $null
$script:credAD = $null
$script:credExchange = $null
$script:credSkype = $null
$script:credNewAccount = $null
$script:strHybrid = $null
$script:strEasPolicy = $null
$script:strDatabase = $null
$status = @{}

###############################################################################
#
#    COMMON FUNCTIONS
#
###############################################################################

function CountDown() {
    param($timeSpan)

    while ($timeSpan -gt 0)
  {
    Write-Host '.' -NoNewline
    $timeSpan = $timeSpan - 1
    Start-Sleep -Seconds 1
  }
}
	
function CleanupAndFail {
  # Cleans up and prints an error message
    
  param
  (
    $strMsg
  )
  if ($strMsg)
    {
        PrintError -strMsg ($strMsg)

    }
    Cleanup
    exit 1
}

function Cleanup () {
  # Cleans up set state such as remote powershell sessions
    if ($sessExchange)
    {
        Remove-PSSession -Id $sessExchange
    }
    if ($sessCS)
    {
        Remove-PSSession -Id $sessSkype
    }
}

function PrintError {
    
  param
  (
    $strMsg
  )
  Write-Host $strMsg
}

function PrintSuccess {
    
  param
  (
    $strMsg
  )
  Write-Host $strMsg
}

function PrintAction {
    
  param
  (
    $strMsg
  )
  Write-Host $strMsg
}

function ExitIfError {
    
  param
  (
    $strMsg
  )
  if ($Error)
    {
        CleanupAndFail -strMsg ($strMsg)

    }
}

function DisplayIntroduction {
  #Welcome Screen
  Clear-Host
  Write-Host '*********************************************************************'
  Write-Host ' Microsoft Surface Hub and Skype Room Systems v2 Provisioning Script '
  Write-Host '*********************************************************************'
  Write-Host ''
}

function DeviceChoice {
  Write-Host 'Welcome.'
  Write-Host 'Do you want to provision a'
  Write-Host 'Microsoft Surface Hub or a Skype Room Systems v2 device.'
  $script:strDevice = Read-Host -Prompt 'Enter 1 for a Surface Hub or 2 for a Skype Room Systems v2 device. Or enter 9 to exit and do nothing.'
  if ($script:strDevice -eq 1)
    {
      SHRoutine
    }
    else
    {	
      if ($script:strDevice -eq 2)
        {
          SRSRoutine
        }
        else
        {
          if ($script:strDevice -eq 9)
            {
              Terminate
            }
            else
            {
              Clear-Host
              Write-Host 'Sorry, invalid entry'
              Start-Sleep -Seconds 1
              Clear-Host
              DeviceChoice
            }
        }
    }
}

function SHRoutine {
  Clear-Host
  Write-Host '*********************'
  Write-Host 'Microsoft Surface Hub'
  Write-Host '*********************'
  Write-Host 'You can create just the AD account(s), Exchange account(s),'
  Write-Host 'Skype account(s), or all of the above.'
  $strProvisionMode = Read-Host -Prompt 'Enter 1 for ALL, 2 for AD only, 3 for Exchange only or 4 for Skype only'
  if ($strProvisionMode -eq 1)
    {
      ProvisionAll
    }
    else
    {	
      if ($strProvisionMode -eq 2)
        {
          ProvisionAD
        }
        else
        {
          if ($strProvisionMode -eq 3)
            {
              ProvisionExchange
            }
            else
            {
              if ($strProvisionMode -eq 4)
                {
                  ProvisionSkype
                }
                else
                {
                  Clear-Host
                  Write-Host 'Sorry, invalid entry'
                  Start-Sleep -Seconds 1
                  Clear-Host
                  SHRoutine
                }
            }
        }
    }
		
	
}

function SRSRoutine {
  Clear-Host
  Write-Host '*********************'
  Write-Host 'Skype Room Systems v2'
  Write-Host '*********************'
  Write-Host 'You can Create just the AD account(s), Exchange account(s),'
  Write-Host 'Skype account(s), or all of the above.'
  $strProvisionMode = Read-Host -Prompt 'Enter 1 for ALL, 2 for AD only, 3 for Exchange only or 4 for Skype only'
  if ($strProvisionMode -eq 1)
    {
      ProvisionAll
    }
    else
    {	
      if ($strProvisionMode -eq 2)
        {
          ProvisionAD
        }
        else
        {
          if ($strProvisionMode -eq 3)
            {
              ProvisionExchange
            }
            else
            {
              if ($strProvisionMode -eq 4)
                {
                  ProvisionSkype
                }
                else
                {
                  Clear-Host
                  Write-Host 'Sorry, invalid entry'
                  Start-Sleep -Seconds 1
                  Clear-Host
                  SRSRoutine
                }
            }
        }
    }
}

function ProvisionAll {
  ProvisionAD
  ProvisionExchange
  ProvisionSkype
}

function ProvisionAD {
  Clear-Host
  Write-Host '*************************************'
  Write-Host 'Provision Active Directory / Azure AD'
  Write-Host '*************************************'
  Write-Host 'Will you be creating your AD user on premises or in the cloud?'
  $strADUser = Read-Host -Prompt 'Enter 1 for on premises or 2 for in the cloud.'
  if ($strADUser -eq 1)
    {
      ProvisionLocalAD
    }
    else
    {
      if ($strADUser -eq 2)
        {
          ProvisionCloudAD
        }
        else
        {
          Clear-Host
          Write-Host 'Sorry, invalid entry'
          Start-Sleep -Seconds 1
          Clear-Host
          ProvisionAD
        }
    }
}

function ProvisionLocalAD {
  Clear-Host
  Write-Host '********************************'
  Write-Host 'Provision Local Active Directory'
  Write-Host '********************************'
  GatherHybridState
  GatherLocalADCreds
  CreateLocalAD
	
  if ($strHybrid -eq 1 -and $strProvisionMode -eq 1)
  {
    Write-Host 'Please wait for an AADSync to occur and validate the user has been replicated.'
    Write-Host 'Once you ensure the account exists in Office 365 please continue.'
    Pause
  }
}

function GatherHybridState {
  $script:strHybrid = Read-Host -Prompt 'Are the users in your tenant synced to Office 365? 1 for Yes 2 for No'
  if (!($strHybrid -eq 1 -or $strHybrid -eq 2))
    {
      Clear-Host
      Write-Host 'Sorry, invalid entry'
      Start-Sleep -Seconds 1
      Clear-Host
      GatherHybridState
    }

}

function GatherLocalADCreds () {
  #	$credAD = $null
  $script:credAD = Get-Credential -Message 'Enter credentials of an Domain Admin or a user with account creation rights'
  if (!$credAD)
  {
      CleanupAndFail -strMsg ('Valid credentials are required to create and prepare the account.')

  }
}

function CreateLocalAD () {
  ## Collect account data ##
  $script:credNewAccount = (Get-Credential -Message 'Enter the desired UPN and password for this new account')
  $strUpn = $credNewAccount.UserName
  $strAlias = $credNewAccount.UserName.substring(0,$credNewAccount.UserName.indexOf('@'))
  $strDisplayName = Read-Host -Prompt "Please enter the display name you would like to use for $strUpn"1
  if (!$credNewAccount -Or [string]::IsNullOrEmpty($strDisplayName) -Or [string]::IsNullOrEmpty($credNewAccount.UserName) -Or $credNewAccount.Password.Length -le 0)
  {
      CleanupAndFail -strMsg 'Please enter all of the requested data to continue.'
      exit 1
  }
    if ($strProvisionMode -eq 1 -or $strProvisionMode -eq 2)
    {
    New-ADUser -UserPrincipalName $credNewAccount.UserName -SamAccountName $strAlias -AccountPassword $credNewAccount.Password -Name $strDisplayName -DisplayName $strDisplayName -PasswordNeverExpires $true -CannotChangePassword $true -Enabled $false
    }
    Countdown -timeSpan 30
}

function CaptureRoomUPN {
  ## Collect account data ##
  $script:credNewAccount = (Get-Credential -Message 'Enter the existing UPN and password for the room account')
}

function ProvisionCloudAD {
  Clear-Host
  Write-Host '********************************'
  Write-Host 'Provision Azure Active Directory'
  Write-Host '********************************'
  GatherHybridState
  GatherCloudADCreds
  Connect2AzureAD
  CreateCloudAD
  AssignPrimaryCloudLicenses
	
  if ($strHybrid -eq 1 -and $strProvisionMode -eq 1)
  {
    Write-Host 'Please wait for an AADSync to occur and vaidate the user has been replicated.'
    Write-Host 'Once you ensure the account exists in Office 365 the script will continue.'
    Pause
  }
}

function GatherCloudADCreds () {
  #	$credAD = $null
  $script:credAD = Get-Credential -Message 'Please enter the credentials of an O365 Global Administrator'
  if (!$credAD)
  {
      CleanupAndFail -strMsg ('Valid credentials are required to create and prepare the room account.')

  }
}

function Connect2AzureAD {
  try
  {
      Connect-MsolService -Credential $credAD
  }
  catch
  {
      CleanupAndFail -strMsg "Failed to connect to Azure Active Directory. Please check your credentials and try again. Error message: $_"
  }
}

function CreateCloudAD {
  ## Collect account data ##
  $script:credNewAccount = (Get-Credential -Message 'Enter the desired UPN and password for this new room account')
  $strUpn = $credNewAccount.UserName
  $strAlias = $credNewAccount.UserName.substring(0,$credNewAccount.UserName.indexOf('@'))
  $strDisplayName = Read-Host -Prompt "Please enter the display name you would like to use for $strUpn"
  if (!$credNewAccount -Or [string]::IsNullOrEmpty($strDisplayName) -Or [string]::IsNullOrEmpty($credNewAccount.UserName) -Or $credNewAccount.Password.Length -le 7)
  {
    CleanupAndFail -strMsg 'Please enter all of the requested data to continue.'
    exit 1
  }
  if ($strProvisionMode -eq 1 -or $strProvisionMode -eq 2)
  {
    try 
    {
      $Error.Clear()
      $strPlainPass = $credNewAccount.GetNetworkCredential().Password
      New-MsolUser -UserPrincipalName $credNewAccount.UserName -DisplayName $strDisplayName -Password $strPlainPass -PasswordNeverExpires $True -ForceChangePassword $false -UsageLocation US
    }
    catch
    {
    }
    if ($Error)
    {
      $Error.Clear()
      $status['Azure Account Create'] = 'Failed to create Azure AD room account. Please validate if you have the appropriate access.'
    }
          else
      {
        $status['Azure Account Create'] = "Successfully added $strDisplayName to Azure AD"
      }
  }
}

function AssignPrimaryCloudLicenses {
  Clear-Host
    PrintAction -strMsg 'We found the following licenses available for your tenant:'
    $skus = (Get-MsolAccountSku | Where-Object { !$_.AccountSkuID.Contains('INTUNE') -and !$_.AccountSkuID.Contains('PSTN')})
    $i = 1
    Foreach ($strSKU in $skus)
  {
    Write-Host -NoNewline $i
    Write-Host -NoNewLine ': AccountSKUID: '
    Write-Host -NoNewLine $strSKU.AccountSkuid
    Write-Host -NoNewLine ' Active Units: '
    Write-Host -NoNewLine $strSKU.ActiveUnits
    Write-Host -NoNewLine ' Unassigned Units: '
    $iUnassigned = $strSKU.ActiveUnits - $strSKU.ConsumedUnits
    Write-Host $iUnassigned
    $i++
  }
	
  $iLicenseIndex = 0

      do
      {
        $iLicenseIndex = Read-Host -Prompt 'Choose the number for the SKU you want to assign to the room account'
      }
		
    while ($iLicenseIndex -lt 1 -or $iLicenseIndex -gt $skus.Length)
      $strLicenses = $skus[$iLicenseIndex - 1].AccountSkuId
    $strSkuPartNumber = $skus[$iLicenseIndex - 1].SkuPartNumber

      if (![string]::IsNullOrEmpty($strLicenses))
      {
          try 
          {
              $Error.Clear()
              Set-MsolUserLicense -UserPrincipalName $credNewAccount.UserName -AddLicenses $strLicenses
          }
          catch
          {
          }
			
          if ($Error)
          {
              $Error.Clear()
              $status['Office 365 License'] = 'Failed to add a license to the room account. Please make verify if you have enough unassigned licenses.'
          }
          else
            {
                $status['Office 365 License'] = "Successfully added $strSkuPartNumber the account"
          $strMoreLicenses = Read-Host -Prompt 'Would you like to add another SKU such as PSTN Calling? Enter 1 for Yes or 2 for No'
          if ($strMoreLicenses -eq 1)
          {
            AssignOtherCloudLicenses
          }	
            else
            {
              if (!($strMoreLicenses -eq 2))
              {
                Clear-Host
                Write-Host 'Sorry, invalid entry'
                Start-Sleep -Seconds 1
                Clear-Host
                ProvisionAD
                }
            }
            }
    }
}

function AssignOtherCloudLicenses {

    PrintAction -strMsg 'We found the following licenses available for your tenant:'
    $skus = (Get-MsolAccountSku | Where-Object { !$_.AccountSkuID.Contains('INTUNE')
  })
    $i = 1
    Foreach ($strSKU in $skus)
  {
    $iUnassigned = $strSKU.ActiveUnits - $strSKU.ConsumedUnits
    	
	  Write-Host -NoNewLine $i
	  Write-Host -NoNewLine ': AccountSKUID: '
	  Write-Host -NoNewLine $strSKU.AccountSkuID
	  Write-Host -NoNewLine ' Acctive Units: '
	  Write-Host -NoNewLine $strSKU.ActiveUnits
	  Write-Host -NoNewLine ' Unassigned Units '
	  Write-Host $iUnassigned
	
	
    $i++
  }
	
  $iLicenseIndex = 0

      do
      {
        $iLicenseIndex = Read-Host -Prompt 'Choose the number for the SKU you want to add'
      }
		
    while ($iLicenseIndex -lt 1 -or $iLicenseIndex -gt $skus.Length)
      $strLicenses = $skus[$iLicenseIndex - 1].AccountSkuId
    $strSkuPartNumber = $skus[$iLicenseIndex - 1].SkuPartNumber

      if (![string]::IsNullOrEmpty($strLicenses))
      {
          try 
          {
              $Error.Clear()
              Set-MsolUserLicense -UserPrincipalName $credNewAccount.UserName -AddLicenses $strLicenses
          }
          catch
          {
          }
			
          if ($Error)
          {
              $Error.Clear()
              $status['Office 365 License'] = 'Failed to add the additional license to the room account. Please verify if you have enough unassigned licenses.'
          }
          else
            {
                $status['Office 365 Additional License'] = "Successfully added $strSkuPartNumber to the account"
          $strMoreLicenses = Read-Host -Prompt 'Would you like to add any other SKU to the room account? Enter 1 for Yes or 2 for No'
          if ($strMoreLicenses -eq 1)
          {
            AssignOtherCloudLicenses
          }	
            else
            {
              if (!($strMoreLicenses -eq 2))
              {
                Clear-Host
                Write-Host 'Sorry, invalid entry'
                Start-Sleep -Seconds 1
                Clear-Host
                ProvisionAD
                }
            }

            }
    }
}

function ProvisionExchange {
  if (!$credNewAccount)
  {
    CaptureRoomUPN
  }
  Clear-Host
  Write-Host '********************************'
  Write-Host 'Provision Exchange Room Mailbox'
  Write-Host '********************************'
  Write-Host 'Will you be creating your Exchange user on premises or in the cloud?'
  $strExchangeUser = Read-Host -Prompt 'Enter 1 for on premises / hybrid or 2 for in the cloud.'
  if ($strExchangeUser -eq 1)
    {
      ProvisionLocalExchange
    }
    else
    {
      if ($strExchangeUser -eq 2)
        {
          ProvisionCloudExchange
        }
        else
        {
          Clear-Host
          Write-Host 'Sorry, invalid entry'
          Start-Sleep -Seconds 1
          Clear-Host
          ProvisionExchange
        }
    }
}

function ProvisionLocalExchange {
  Clear-Host
  Write-Host '*************************************'
  Write-Host 'Provision Local Exchange Room Mailbox'
  Write-Host '*************************************'
  if (!$credAD)
    {
      CaptureExchangeCreds
    }
    else
    {
      Write-Host 'Are the Exchange user with mailbox creation rights different than the AD credentials?'
      $strSameCreds = Read-Host -Prompt 'Enter 1 for Yes and 2 for No'
      if ($strSameCreds -eq 1)
      {
        CaptureExchangeCreds
      }
      else
      {
        if ($strSameCreds -eq 2)
        {
          $script:credExchange = $credAD
        }
        else
        {
          Clear-Host
          Write-Host 'Sorry, invalid entry'
          Start-Sleep -Seconds 1
          Clear-Host
          ProvisionLocalExchange
        }
      }

    }
	
  $strExchangeServer = Read-Host -Prompt 'Please enter the FQDN of your exchange server (e.g. exch.contoso.com)'
  PrintAction -strMsg 'Connecting to remote sessions. This can occasionally take a while - please do not enter input...'
	
  try 
  {
      $sessExchange = New-PSSession -ConfigurationName microsoft.exchange -Credential $credExchange -AllowRedirection -Authentication Kerberos -ConnectionUri "http://$strExchangeServer/powershell" -WarningAction SilentlyContinue
  }
	
  catch
  {
      CleanupAndFail -strMsg ("Failed to connect to exchange. Please check your credentials and try again. If this continues to fail, you may not have permission for remote powershell - if not, please perform the setup manually. Error message: $_")
  }
	
  PrintSuccess -strMsg 'Connected to Remote Exchange Shell'
  Import-PSSession -Session $sessExchange -AllowClobber -WarningAction SilentlyContinue
	
  # In case there was any uncaught errors
  # ExitIfError -strMsg ('Remote connections failed. Please check your credentials and try again.')
	
  ExchangeDBs
  if($script:strDevice -eq 1)
  {
    EASPolicy
  }
		
  $Error.Clear()
  PrintAction -strMsg 'Creating a new account...'
	
  try
  {
      Enable-Mailbox $credNewAccount.UserName -Alias $credNewAccount.UserName.substring(0,$credNewAccount.UserName.indexOf('@')) -Database $strDatabase
      Clear-Host
    PrintAction -strMsg 'Creating Mailbox'
    Countdown -timeSpan 30
    if($script:strDevice -eq 1) {Set-CASMailbox $credNewAccount.UserName -ActiveSyncMailboxPolicy $strEASpolicy}
    Set-Mailbox $credNewAccount.UserName -Type Room -EnableRoomMailboxAccount $true -RoomMailboxPassword $credNewAccount.Password
    #Need to capture error and give direction
    #You don't have permission to directly change a mailbox account password without providing old password. Reset Password role is required for directly changing password.
  }
	
  catch
  {
  }
	
  ExitIfError -strMsg 'Failed to create a new mailbox on exchange.'

	
  $status['Mailbox Setup'] = 'Successfully created a mailbox for the new account'
  Clear-Host
  $mailbox = Get-Mailbox $credNewAccount.UserName
  $strEmail = $mailbox.WindowsEmailAddress
  PrintSuccess -strMsg "The following mailbox has been created for this room: $strEmail"
  ExchMBXProps
}

function ProvisionCloudExchange {
  Clear-Host
  Write-Host '*************************************'
  Write-Host 'Provision O365 Exchange Room Mailbox'
  Write-Host '*************************************'
  if (!$credAD)
    {
      CaptureEXOCreds
    }
    else
    {
      Write-Host 'Is the Exchange Online admin account different than the Global Administrator account?'
      $strSameCreds = Read-Host -Prompt 'Enter 1 for Yes and 2 for No'
      if ($strSameCreds -eq 1)
      {
        CaptureEXOCreds
        Countdown -timeSpan 120
      }
      else
      {
        if ($strSameCreds -eq 2)
        {
          $script:credExchange = $credAD
          Write-Host 'Waiting for Exchange Online Mailbox to be provisioned'
          Countdown -timeSpan 120
        }
        else
        {
          Clear-Host
          Write-Host 'Sorry, invalid entry'
          Start-Sleep -Seconds 1
          Clear-Host
          ProvisionCloudExchange
        }
      }

    }
  Connect2EXO
  if ($script:strDevice -eq 1) { EASPolicy }
  $Error.Clear()
  PrintAction -strMsg 'Converting user mailbox to a room account...'  
  PrintAction -strMsg 'This process will take around 10 minutes to sync the newly created mailbox to O365 servers.'  
  PrintAction -strMsg 'Please wait.' 
  Countdown -timeSpan 600
  try
  {
      $mailbox = $null
      $mailbox = (Set-Mailbox -Identity $credNewAccount.UserName -MicrosoftOnlineServicesID $credNewAccount.UserName -Type Room)
      $mailbox = Get-Mailbox -Identity $credNewAccount.UserName
    if ($script:strDevice -eq 1) { Set-CASMailbox $credNewAccount.UserName -ActiveSyncMailboxPolicy $strEASpolicy }
    Countdown -timeSpan 120
		
  }

  catch
  {
    PrintAction -strMsg 'Converting mailbox to a room account failed you can do this step manually.' 
  
  }
  ExitIfError -strMsg 'Failed to configure a new room mailbox on exchange.'
  
  $status['Mailbox Setup'] = "Successfully configured a room mailbox for the account $($credNewAccount.UserName)"

  $strEmail = $mailbox.WindowsEmailAddress
  Clear-Host
  PrintSuccess -strMsg "The following mailbox has been created for this room: $strEmail"

  ExchMBXProps
}

function Connect2EXO {
  try
  {
    $sessEXO = New-PSSession -ConfigurationName microsoft.exchange -Credential $credExchange -AllowRedirection -Authentication basic -ConnectionUri 'https://outlook.office365.com/powershell-liveid/' -WarningAction SilentlyContinue
    Import-PSSession -Session $sessEXO -AllowClobber -WarningAction SilentlyContinue
  }
  catch
  {
    CleanupAndFail -strMsg "Failed to connect to Exchange Online. Please check your credentials and try again. Error message: $_"
  }
}

function CaptureExchangeCreds {
  $script:credExchange = Get-Credential -Message 'Enter credentials of an Exchange user with mailbox creation rights'
  if (!$credExchange)
  {
      CleanupAndFail -strMsg ('Valid credentials are required to create and prepare the account.')

  }
}

function CaptureEXOCreds () {
  #	$credAD = $null
  $script:credExchange = Get-Credential -Message 'Enter credentials of an Exchange Online Administrator'
  if (!$credExchange)
  {
      CleanupAndFail -strMsg ('Valid credentials are required to create and prepare the account.')

  }
}

function ExchangeDBs {
  Clear-Host
  PrintAction -strMsg 'We found the following databases in you on premises Exchange environment:'
  $exDBs = (Get-MailboxDatabase)
  $i = 1
  Foreach ($DB in $exDBs)
  {
    Write-Host -NoNewline $i
    Write-Host -NoNewLine ': Database Name: '
    Write-Host $DB.Name
    $i++
  }
	
    $iDatabaseIndex = 0

    do

  {
        $iDatabaseIndex = Read-Host -Prompt 'Choose the database you want the Room Mailbox created in'
    }
  while ($iDatabaseIndex -lt 1 -or $iDatabaseIndex -gt $exDBs.Length)
    $script:strDatabase = $exDBs[$iDatabaseIndex - 1].Name
}

function EASPolicy {
  Clear-Host
  PrintAction -strMsg 'A valid Exchange Active Sync Policy is Required. Would you like to use an existing Policy or Create a new one?'
  $strEASDecision = Read-Host -Prompt 'Enter 1 for Existing or 2 for New'
  if ($strEASDecision -eq 1)
    {
    EASPolicyExisting
    }
    else
    {
      if ($strEASDecision -eq 2)
        {
        EASPolicyNew
        }
        else
        {
          Clear-Host
          Write-Host 'Sorry, invalid entry'
          Start-Sleep -Seconds 1
          Clear-Host
          ProvisionExchange
        }
    }
					
}

function EASPolicyExisting {
  PrintAction -strMsg 'A valid Exchange Active Sync Policy is Required - the following policies were found:'
  $exEasPol = (Get-MobileDeviceMailboxPolicy)
  $i = 1
  Foreach ($POL in $exEasPol)
  {
    if ($POL.PasswordEnabled -eq $false)
    {
      $strValidity = 'Valid, will work'
    }
    else
    {
      $strValidity = 'Password Policy Enabled, will not work'
    }
    Write-Host -NoNewline $i
    Write-Host -NoNewLine ': Policy Name: '
    Write-Host -NoNewline $POL.Name
    Write-Host -NoNewline ' - '
    Write-Host $strValidity
    $i++
  }
	
    $iPolicyIndex = 0

    do
  {
        $iPolicyIndex = Read-Host -Prompt 'Choose the Exchange Policy you would like to use:'
    }

  while ($iPolicyIndex -lt 1 -or $iPolicyIndex -gt $exEasPol.Length)
    $script:strEasPolicy = $exEasPol[$iPolicyIndex - 1].Name
}

function EASPolicyNew {
  $script:strEasPolicy = (New-MobileDeviceMailboxPolicy -Name 'SurfaceHubs' -PasswordEnabled $false)
}

function ExchMBXProps {
  PrintAction -strMsg 'You may optionally configure Mailbox Room Properties using this script (or manually at a later time)'
  PrintAction -strMsg 'Microsoft recommends enabling Automatic Calendar Processing with a response - would you like to configure'
  PrintAction -strMsg 'these options now via the script?'
  $strMBXPropsDecision = Read-Host -Prompt 'Enter 1 for Yes or 2 for No'
  if ($strMBXPropsDecision -eq 1)
    {
  
      Set-CalendarProcessing -Identity $credNewAccount.UserName -AutomateProcessing AutoAccept -AddOrganizerToSubject $false -AllowConflicts $false -DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false
      Set-CalendarProcessing -Identity $credNewAccount.UserName -AddAdditionalResponse $true -AdditionalResponse 'This is a <tla rid="Skype Meeting"/> capable room!'
    }
    else
    {
      if (!($strMBXPropsDecision -eq 2))
        {
          Clear-Host
          Write-Host 'Sorry, invalid entry'
          Start-Sleep -Seconds 1
          Clear-Host
          ProvisionExchange
        }
    }
}

function ProvisionSkype {
  if (!$credNewAccount)
  {
    CaptureRoomUPN
  }
  Clear-Host
  Write-Host '********************************'
  Write-Host 'Provision Skype'
  Write-Host '********************************'
  Write-Host 'Will you be creating your Skype user on premises or in the cloud?'
  $strSkypeUser = Read-Host -Prompt 'Enter 1 for on premises / hybrid or 2 for in the cloud.'
  if ($strSkypeUser -eq 1)
    {
      ProvisionLocalSkype
    }
    else
    {
      if ($strSkypeUser -eq 2)
        {
          ProvisionCloudSkype
        }
        else
        {
          Clear-Host
          Write-Host 'Sorry, invalid entry'
          Start-Sleep -Seconds 1
          Clear-Host
          ProvisionSkype
        }
    }
}

function ProvisionLocalSkype {
  Clear-Host
  Write-Host '*************************************'
  Write-Host 'Provision Local Skyp User'
  Write-Host '*************************************'
  if (!$credAD)
    {
      CaptureSkypeCreds
    }
    else
    {
      Write-Host 'Are the Skype Administrator credentials different than the AD credentials?'
      $strSameCreds = Read-Host -Prompt 'Enter 1 for Yes and 2 for No'
      if ($strSameCreds -eq 1)
      {
        CaptureSkypeCreds
      }
      else
      {
        if ($strSameCreds -eq 2)
        {
          $script:credSkype = $credAD
        }
        else
        {
          Clear-Host
          Write-Host 'Sorry, invalid entry'
          Start-Sleep -Seconds 1
          Clear-Host
          ProvisionLocalExchange
        }
      }

    }
	
  $strSkypeFQDN = Read-Host -Prompt 'Please enter the FQDN of a Skype server (e.g. skype.contoso.com) for Remote PowerShell'
  PrintAction -strMsg 'Connecting to remote session. This can occasionally take a while - please do not enter input...'
	
  try 
  {
      $sessSkype = New-PSSession -Credential $credSkype -ConnectionURI "https://$strSkypeFQDN/OcsPowershell" -AllowRedirection -WarningAction SilentlyContinue
  }
	
  catch
  {
      CleanupAndFail -strMsg ("Failed to connect to exchange. Please check your credentials and try again. If this continues to fail, you may not have permission for remote powershell - if not, please perform the setup manually. Error message: $_")
  }
	
  PrintSuccess -strMsg 'Connected to Remote Skype Shell'
  Import-PSSession -Session $sessSkype -AllowClobber -WarningAction SilentlyContinue
	
  # In case there was any uncaught errors
  ExitIfError -strMsg ('Remote connections failed. Please check your credentials and try again.')
	
  $Error.Clear()
  PrintAction -strMsg 'Configuring account for Skype for Business.'
	
  # Getting registrar pool
  Clear-Host
  $strRegPool = $strSkypeFQDN
  $Error.Clear()
  $strRegPoolEntry = Read-Host -Prompt "Enter a Skype for Business Registrar Pool, or leave blank if [$strRegPool] is a Standard Editon Pool"
	
  if (![string]::IsNullOrEmpty($strRegPoolEntry))
  {
      $strRegPool = $strRegPoolEntry
  }
	
  PrintAction -strMsg 'Enabling Skype for Business...'
  $Error.Clear()

	
  try
  {
      Enable-CsMeetingRoom -Identity $credNewAccount.UserName -RegistrarPool $strRegPool -SipAddressType EmailAddress
    Countdown -timeSpan 30
  }
	
  catch
  {
  } 

  if ($Error)
  {
      PrintError -strMsg 'Failed to setup the Skype for Business meeting room resource - you can run this Skype Room Provisioning Script to try again.'
      $Error.Clear()

  }
  else
  {
      PrintSuccess -strMsg 'Successfully enabled the account as a Skype for Business meeting room'
    ProvisionSkypeEV
  }
	
}

function ProvisionCloudSkype {
  Clear-Host
  Write-Host '*******************************************'
  Write-Host 'Provision Skype for Business Online Account'
  Write-Host '*******************************************'
  if (!$credAD)
    {
      CaptureS4BOCreds
    }
    else
    {
      Write-Host 'Is the Skype for Business Online admin account different than the Global Administrator account?'
      $strSameCreds = Read-Host -Prompt 'Enter 1 for Yes and 2 for No'
      if ($strSameCreds -eq 1)
      {
        CaptureS4BOCreds
      }
      else
      {
        if ($strSameCreds -eq 2)
        {
          $script:credSkype = $credAD
        }
        else
        {
          Clear-Host
          Write-Host 'Sorry, invalid entry'
          Start-Sleep -Seconds 1
          Clear-Host
          ProvisionCloudExchange
        }
      }

    }
  Connect2S4BO
  # Getting registrar pool
	
    try
      {
       PrintAction -strMsg 'We have now connected to Skype for Business, please wait while we attempt to set up the meeting room' 
       Countdown -timeSpan 300
 
        $strRegPool = $null

        $strRegPool = (Get-CsTenant).TenantPoolExtension
    }
	
    catch
    {
    }
    
  if ($Error)
    {
        $Error.Clear()

        $strRegPool = ''

        Write-Host 'We failed to lookup your Skype for Business Registrar Pool, but you can still enter it manually'
    }
	
    else
    {
        $strRegPool = $strRegPool[0].Substring($strRegPool[0].IndexOf(':') + 1)
    }
	
  $Error.Clear()
  try
  {
    Enable-CsMeetingRoom -Identity $credNewAccount.UserName -RegistrarPool $strRegPool -SipAddressType EmailAddress
  }
	
  catch
  {
  }

  ExitIfError -strMsg ('Failed to setup Skype for Business meeting room')

  PrintSuccess -strMsg "Successfully enabled $strRoomUri as a Skype for Business meeting room"
  ProvisionSkypeEV
}

function Connect2S4BO {
  try
  {
    $sessSkype = New-CsOnlineSession -Credential $credSkype
    Import-PSSession -Session $sessSkype -AllowClobber -WarningAction SilentlyContinue
  }
  catch
  {
    CleanupAndFail -strMsg "Failed to connect to Skype for Business Online. Please check your credentials and try again. Error message: $_"
  }
}

function ProvisionSkypeEV {
  Clear-Host
  PrintAction -strMsg 'Would you like to configure Enterprise Voice for the Room?'
  $strSkypeEVDecision = Read-Host -Prompt 'Enter 1 for Yes or 2 for No'
    if ($strSkypeEVDecision -eq 1)
    {
      PrintAction -strMsg 'A valid E.164 number is required (ex. 15255551000 or 1525551000;ext=1000)'
      $strEVInput = Read-Host -Prompt 'Enter a valid E.164 number'
      if (!$strEVInput)
        {
          Clear-Host
          Write-Host 'Sorry, invalid entry'
          Start-Sleep -Seconds 1
          Clear-Host
          ProvisionSkypeEV
        }
        else
        {
          try
          {
            $LineURI = 'tel:+' + $strEVInput
            PrintAction -strMsg "Your LineURI will be configured as $LineURI - is that correct?"
            $strEVOkay = Read-Host -Prompt 'Enter 1 for Accept and configure or 2 to Go Back'
            if ($strEVOkay -eq 1)
              {
                Set-CsMeetingRoom -Identity $credNewAccount.UserName -LineURI $LineURI -EnterpriseVoiceEnabled $true
              }
              else
              {
              if ($strEVOkay -eq 2)
                {
                  ProvisionSkypeEV
                }
                else
                {
                  Clear-Host
                  Write-Host 'Sorry, invalid entry'
                  Start-Sleep -Seconds 1
                  Clear-Host
                  ProvisionSkypeEV
                }
              }
          }
          catch
          {
          }
        }
					
    }
    else
    {
      if (!($strSkypeEVDecision -eq 2))
        {
          Clear-Host
          Write-Host 'Sorry, invalid entry'
          Start-Sleep -Seconds 1
          Clear-Host
          ProvisionExchange
        }
    }
}

function CaptureS4BOCreds () {
  #	$credAD = $null
  $script:credSkype = Get-Credential -Message 'Enter credentials of an Skype for Business Online Administrator'
  if (!$credSkype)
  {
      CleanupAndFail -strMsg ('Valid credentials are required to create and prepare the account.')

  }
}

function Terminate {
  Clear-Host
  Write-Host 'Application Terminated as Requested'
  Start-Sleep -Seconds 1
  Exit 1
}

DisplayIntroduction
DeviceChoice
Cleanup
Clear-Host
PrintAction -strMsg 'Summary for Actions'
if ($status.Count -gt 0)
{
    ForEach($k in $status.Keys) 
    {
        $v = $status[$k]
        $color = 'yellow'
        if ($v[0] -eq 'S') { $color = 'green' }
        elseif ($v[0] -eq 'F') 
        {
            $color = 'red' 
            $v += ' Go to http://aka.ms/shubtshoot for help'
        }

        Write-Host $v
    }
}
else
{
    PrintError -strMsg 'The process could not be completed'
}
