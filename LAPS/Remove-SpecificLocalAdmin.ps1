$computerName = hostname
$LocalGroupName = "Administrators"

$Group = [ADSI]("WinNT://$computerName/$localGroupName,group")
$Group.Members() | foreach {

    $AdsPath = $_.GetType().InvokeMember('Adspath', 'GetProperty', $null, $_, $null)
    $A = $AdsPath.split('/',[StringSplitOptions]::RemoveEmptyEntries)
    $Names = $a[-1]
    $Domain = $a[-2]

    If ($Domain -eq $ComputerName -and $Names -ne "Administrator") {
        Add-Content C:\Windows\Temp\RemoveUsersFromAdminGroup.log "User $Names found on computer $computerName … "
        if ($Names -eq "bdtmsdlaps") {
		net user bdtmsdlaps /DELETE
		Add-Content C:\Windows\Temp\RemoveUsersFromAdminGroup.log "Removed"
	}

    }

}
