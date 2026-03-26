<# 
Automatic Proxy PAC file creator for Office 365 Address space
based on Office 365 XML feed:

https://support.content.office.net/en-us/static/O365IPAddresses.xml

THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY 
OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE 
IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR RESULTS FROM THE USE OF 
THIS CODE REMAINS WITH THE USER.

Author:		Aaron Guilmette
			aaron.guilmette@microsoft.com
			
Find the most updated version at:
https://gallery.technet.microsoft.com/Office-365-Proxy-Pac-60fb28f7
#>

<#
.SYNOPSIS
Automatically generate a Proxy Automatic Configuration (PAC) file
from the Office 365 URL and IP addresses XML feed.

.PARAMETER AlwaysProxyBlockList
Add the domains in Blocklist to a "PROXY ONLY" directive with no fail over to
DIRECT.

.PARAMETER Blocklist
Use the "blocklist" to exclude domains and patterns from the proxy bypass list.
Do not use wildcards, as RegEx matching can be used to filter entries out.  If
AlwaysProxyBlockList is used, these entries will get added with a PROXY only
directive.

.PARAMETER GenerateIPList
Use to generate a text file containing all IP address ranges for selected products.

.PARAMETER ImportFile
Use specified XML import file instead of downloading from support site.

.PARAMETER IncludeIPAddresses
Use to include IP addresses in PAC file output.

.PARAMETER OnlyIPAddresses
Use to create only a PAC file containing IP addresses.

.PARAMETER OutputFile
The OutputFile parameter specifies the name of the output PAC file.

.PARAMETER Products
Use the Products parameter to specify which products will be configured in the
PAC. The full list of products keywords that can be used:
	- 'O365' - Office 365 Portal and Shared
	- 'LYO' - Skype for Business (formerly Lync Online)
	- 'Planner' - Planner
	- 'ProPlus' - Office 365 ProPlus
	- 'OneNote' - OneNote
	- 'WAC' - SharePoint WebApps
	- 'Yammer' - Yammer
	- 'EXO' - Exchange online
	- 'Identity' - Office 365 Identity
	- 'SPO' - SharePoint Online
	- 'RCA' - Remote Connectivity Analyzer
	- 'Sway' - Sway
	- 'OfficeMobile' - Office Mobile Apps
	- 'Office365Video' - Office 365 Video
	- 'CRLs' - Certificate Revocation Links
	- 'OfficeiPad' - Office for iPad
	- 'EOP' - Exchange Online Protection
	- 'EX-Fed' - Exchange Federation (?)
    - 'Teams' - Microsoft Teams

.PARAMETER ProxyServer
Use the ProxyServer parameter to specify the proxy server URL or IP
address and port combination.

.PARAMETER USDefense
Use the Defense IP Address list.

.EXAMPLE
.\Office365ProxyPac.ps1 -ProxyServer 10.0.0.1:8080
Configure the PAC file to point to the proxy server at 10.0.0.1:8080.

.EXAMPLE
.\Office365ProxyPac.ps1 -ProxyServer 10.0.0.1:8080 -OutputFile Proxy.pac
Configure the PAC file to point to the proxy server at 10.0.0.1:8080
and write the output to the file Proxy.pac.

.EXAMPLE
.\Office365ProxyPac.ps1 -ProxyServer 5.6.7.8:8080 -Products EXO,LYO
Configure the PAC file to point to proxy server at 5.6.7.8:8080 and only include
entries related to Exhange Online and Skype for Business (formerly Lync Online).

.EXAMPLE
.\Office365ProxyPac.ps1 -ProxyServer 5.6.7.8:8080 -Products EXO,OfficeMobile -Blocklist facebook,youtube
Configure the PAC file to point to proxy server at 5.6.7.8:8080 and only include
entries related to Exchange Online and OfficeMobile, excluding URLs that match
patterns 'facebook' and 'youtube'.

.LINK
https://gallery.technet.microsoft.com/Office-365-Proxy-Pac-60fb28f7

.LINK
https://blogs.technet.microsoft.com/undocumentedfeatures/tag/proxy-automatic-configuration/

.NOTES
2018-01-26	Updated to include US Dept of Defense IP List parameter
2018-01-10	Updated invalid/non-ascii characters
2017-06-06  Updated to include options for IP address inclusion
2017-05-19	Updated to include Skype IP address ranges in proxy bypass list when per PG
#>

[CmdletBinding()]
Param(
	[Parameter(Mandatory=$false,HelpMessage='Always proxy the block list')]
		[switch]$AlwaysProxyBlockList,
	
	[Parameter(Mandatory=$false,HelpMessage='Blocklist in the form of comma-separated domains or patterns')]
		[array]$Blocklist,
	
	[Parameter(Mandatory = $false)]
		[switch]$GenerateIPList,
	
	[Parameter(Mandatory=$false)]
		[string]$ImportFile,
	
	[Parameter(Mandatory = $false)]
		[Switch]$IncludeIPAddresses,
	
	[Parameter(Mandatory = $false)]
		[switch]$OnlyIPAddresses,
	
	[Parameter(Mandatory=$false,HelpMessage='ProxyServer in the form of server:port')]
		[string]$ProxyServer = "10.0.0.1:8080",
	
	[Parameter(Mandatory = $false, HelpMessage = 'Use US Defense IP Address List')]
		[switch]$USDefense,
	
	[Parameter(Mandatory=$false,HelpMessage='OutputFile')]
		[string]$OutputFile = "Office365PAC.pac",
	
	[ValidateSet("O365","LYO","Planner","ProPlus","OneNote","WAC","Yammer","EXO","Identity","SPO","RCA","Sway","OfficeMobile","Office365Video","CRLs","OfficeiPad","EOP","EX-Fed","Teams")]
		[array]$Products = ('O365','LYO','Planner','ProPlus','OneNote','WAC','Yammer','EXO','Identity','SPO','RCA','Sway','OfficeMobile','Office365Video','CRLs','OfficeiPad','EOP','EX-Fed','Teams')
	)

Function cidr
{
	[CmdLetBinding()]
	Param (
		[Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
		[Alias("Length")]
		[ValidateRange(0, 32)]
		$MaskLength
	)
	Process
	{
		Return LongToDotted ([Convert]::ToUInt32($(("1" * $MaskLength).PadRight(32, "0")), 2))
	}
}

Function LongToDotted
{
	[CmdLetBinding()]
	Param (
		[Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
		[String]$IPAddress
	)
	Process
	{
		Switch -RegEx ($IPAddress)
		{
			"([01]{8}\.){3}[01]{8}" {
				Return [String]::Join('.', $($IPAddress.Split('.') | ForEach-Object { [Convert]::ToUInt32($_, 2) }))
			}
			"\d" {
				$IPAddress = [UInt32]$IPAddress
				$DottedIP = $(For ($i = 3; $i -gt -1; $i--)
					{
						$Remainder = $IPAddress % [Math]::Pow(256, $i)
						($IPAddress - $Remainder) / [Math]::Pow(256, $i)
						$IPAddress = $Remainder
					})
				Return [String]::Join('.', $DottedIP)
			}
			default
			{
				
			}
		}
	}
}

Write-Host "The PAC file will be generated for the following products:"
Write-Host $Products

[regex]$ProductsRegEx = '(?i)^(' + (($Products |foreach {[regex]::escape($_)}) -join "|") + ')$'
If ($Blocklist)
	{
	[regex]$BlocklistRegEx = '(?i)(' + (($Blocklist |foreach {[regex]::escape($_)}) -join "|") + ')'
	Write-Host Blocklist is $BlocklistRegEx.ToString()
	}

If ($ImportFile)
	{
	[xml]$O365URLData = Get-Content $ImportFile
	}
Else
	{
	If ($USDefense)
	{
		$O365URL = "https://support.content.office.net/en-us/static/O365IPAddresses_USDefense.xml"
	}
	Else
	{
		$O365URL = "https://support.content.office.net/en-us/static/O365IPAddresses.xml"
	}
	
	Write-Host -ForegroundColor Yellow "Downloading latest Office 365 XML data..."
	[xml]$O365URLData = (New-Object System.Net.WebClient).DownloadString($O365URL)
	}

$SelectedProducts = $O365URLData.SelectNodes("//product") | ? { $_.Name -match $ProductsRegEx }

If (Test-Path $OutputFile) { Remove-Item -Force $OutputFile }

Add-Content $OutputFile "//"
Add-Content $OutputFile "//Office 365 Proxy Automatic Configuration File"
Add-Content $OutputFile "//built from Office 365 Automatic Proxy PAC File Creator"
Add-Content $OutputFile "//"
Add-Content $OutputFile "//Author: aaron.guilmette@microsoft.com"
Add-Content $OutputFile "//Link: https://gallery.technet.microsoft.com/Office-365-Proxy-Pac-60fb28f7"
Add-Content $OutputFile "//"

# Create findProxy Function
Add-Content $OutputFile "function FindProxyForURL(url, host)"
Add-Content $OutputFile "{"
Add-Content $OutputFile "if ("

$IPData = @()
$ProxyURLData = @()
$AlwaysProxyURLMatches = @()

#Added 2017-06-06
If ($OnlyIPAddresses)
{
	$SelectedProducts.AddressList | ? { $_.Type -eq "IPv4" } | % { $address = $_; foreach ($a in $address) { $ProxyURLData += $a.address } }
	#$OutputFileIP = "IPAddresses_" + $OutputFile + ".txt"
}
elseif ($IncludeIPAddresses)
{
	$SelectedProducts.AddressList | ? { $_.Type -eq "IPv4" } | % { $address = $_; foreach ($a in $address) { $ProxyURLData += $a.address } }
	$SelectedProducts.AddressList | ? { $_.Type -eq "URL" } | % { $address = $_; foreach ($a in $address) { $ProxyURLData += $a.address } }
}
Else
{
	$SelectedProducts.AddressList | ? { $_.Type -eq "URL" } | % { $address = $_; foreach ($a in $address) { $ProxyURLData += $a.address } }
	If ($Products -contains "LYO")
	{
		$LYO = $SelectedProducts | ? { $_.Name -eq "LYO" }
		$LYO.AddressList | ? { $_.Type -eq "IPv4" } | % { $address = $_; foreach ($a in $address) { $ProxyURLData += $a.address } }
	}
}

If ($GenerateIPList)
	{
		$SelectedProducts.AddressList | ? { $_.Type -eq "IPv4" } | % { $address = $_; foreach ($a in $address) { $IPData += $a.address } }
		$OutputFileIP = "IPAddresses_" + $OutputFile + ".txt"
	}
	
# Build Proxy List
$IPData = $IPData | Sort -Unique
$ProxyURLData = $ProxyURLData | Sort -Unique
Foreach ($url in $ProxyURLData)
	{
	Write-Host $url
	# Check to see if "AlwaysProxyBlockList" switch param is set
	If ($AlwaysProxyBlockList)
		{
		If ($Blocklist -and $url -match $BlocklistRegEx)
			{
			Write-Host -Fore Red "     URL $($URL) is on the Block list and AlwaysProxyBlockList is configured."
			# Add URL to $AlwaysProxyURLMatches and go process the next URL
			$AlwaysProxyURLMatches += $url
			Continue
			}
		}
	If ($Blocklist -and $url -match $BlocklistRegEx)
		{
		Write-Host -Fore Red "     URL $($URL) is on the Block list and will be skipped."
		Continue
		}
	If ($url -match "\*")
		{
		Add-Content $OutputFile "shExpMatch(host, ""$URL"")||"
		}
	If ($url -match "\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b")
		{
			If ($url -match "/")
			{
				$CIDR = $url.Split("/")[1]
				$IPAddr = $url.Split("/")[0]
				$Mask = cidr $cidr
				Add-Content $OutputFile "isInNet(host,""$IPAddr"",""$Mask"")||"
			}
			Else
			{
				Add-Content $OutputFile "shExpMatch(host, ""$URL"")||"
			}
		}
	Else
	{
		Add-Content $OutputFile "dnsDomainIs(host, ""$URL"")||"
	}
}
Add-Content $OutputFile "dnsDomainIs(host, ""office365.com"")"
Add-Content $OutputFile ")"
Add-Content $OutputFile "return ""DIRECT"";"

# If addresses were added to $AlwaysProxyURLMatches, then build a section to
# send to proxy with no DIRECT failover.
If ($AlwaysProxyURLMatches)
	{
	Add-Content $OutputFile "if ("
		Foreach ($url in $AlwaysProxyURLMatches)
			{
			If ($url -match "\*")
				{
				Add-Content $OutputFile "shExpMatch(host, ""$URL"")||"
				}
			Else
				{
				Add-Content $OutputFile "dnsDomainIs(host, ""$URL"")||"
				}
			}
	Add-Content $OutputFile "dnsDomainIs(host, ""msinvalid.invalid"")"
	Add-Content $OutputFile ")"
	Add-Content $OutputFile "return ""PROXY $ProxyServer"";"
	}

# Add final directive
Add-Content $OutputFile "else { return ""PROXY $ProxyServer; DIRECT"";}"
Add-Content $OutputFile "}"

Try {
	Test-Path $OutputFile -ErrorAction SilentlyContinue > $null
	Write-Host -ForegroundColor Yellow "Done! PAC file is $($OutputFile)."
	}
Catch {
	Write-Host -ForegroundColor Red "PAC file not created."
	}
Finally { }

If ($GenerateIPList)
{
	If ($IPData)
	{
		$IPData | Out-File $OutputFileIP -Force
	}
	Write-Host -ForegroundColor Yellow "IP Address list is saved to $($OutputFileIP)."
}