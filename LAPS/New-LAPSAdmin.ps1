# Must run as 64-bit script in Intune

$lapsAdmin = "wrbc-la"

if (!(Get-LocalUser -Name $lapsAdmin -ErrorAction SilentlyContinue)) {
   [securestring]$password = ConvertTo-SecureString -String (New-Guid) -AsPlainText -Force
   New-LocalUser $lapsAdmin -Password $password -Description "LAPS managed local account"
   Add-LocalGroupMember -Group "Administrators" -Member $lapsAdmin
}
