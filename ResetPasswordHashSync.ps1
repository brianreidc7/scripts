<# 
Reset Office 365 Password Hash Sync configuration for a particular connector.

THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT WARRANTY 
OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE 
IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR RESULTS FROM THE USE OF 
THIS CODE REMAINS WITH THE USER.

Author:  	Aaron Guilmette
			aaron.guilmette@microsoft.com
#>

<#
.SYNOPSIS
Reset Password Hash Sync configuration for connectors.  Works with 
AADSync and AADConnect.

.DESCRIPTION
Reset Password Hash Sync configuration for connectors.  Works with 
AADSync and AADConnect.

.LINK
http://blogs.technet.com/b/undocumentedfeatures/archive/2015/11/18/reset-aadsync-or-aadconnect-password-hash-sync-configuration.aspx

.LINK
https://gallery.technet.microsoft.com/Reset-AADSync-or-f8a0ba2a/file/144862/1/ResetPasswordHashSync.ps1
#>

# Load Modules
If (!(Get-Module ADSync -ListAvailable)) 
	{ 
	Write-Host -ForegroundColor Red "You must run this from the computer with AADSync or AADConnect installed."
	Break
	}
ElseIf (!(Get-Module ADSync))
	{
	Import-Module ADSync 
	}

# Get All Connectors
$adConnectors = Get-ADSyncConnector | ? { $_.Type -eq "AD" }
$aadConnectors = Get-ADSyncConnector | ? { $_.Type -eq "Extensible2" -or $_.SubType -like "*Azure*" }

# Build Menus
$adMenu = @{}
$aadMenu = @{}

Write-Host -Fore Yellow "AD Connectors"
For ($i=1;$i -le $adConnectors.Count; $i++)
	{
	Write-Host "$i. $($adConnectors[$i-1].Name)"
	$adMenu.Add($i,($adConnectors[$i-1].Name))
	}

[int]$adConnectorSelection = Read-Host "Enter Source AD Connector"
$adConnector = $adMenu.Item($adConnectorSelection)
Write-Host "`n"

Write-Host -Fore Yellow "Azure AD / Office 365 Connectors"
For ($i=1;$i -le $aadConnectors.Count; $i++)
	{
	Write-Host "$i. $($aadConnectors[$i-1].Name)"
	$aadMenu.Add($i,($aadConnectors[$i-1].Name))
	}

[int]$aadConnectorSelection = Read-Host "Enter Target Office 365 Connector"
$aadConnector = $aadMenu.Item($aadConnectorSelection)

Write-Host "`n"
Write-Host "Selected AD source connector $adConnector."
Write-Host "Selected AAD target connector $aadConnector.`n"

# Disable and re-enable Password Hash Sync on selected connector
Write-Host "Disabling Password Hash Sync."
Set-ADSyncAADPasswordSyncConfiguration -SourceConnector $adConnector -TargetConnector $aadConnector -Enable $false 
Write-Host "`n"

Write-Host "Enabling Password Hash Sync."
Set-ADSyncAADPasswordSyncConfiguration -SourceConnector $adConnector -TargetConnector $aadConnector -Enable $true

Write-Host "`n"
Write-Host -ForegroundColor Cyan "Password Hash Sync Configuration Done!"