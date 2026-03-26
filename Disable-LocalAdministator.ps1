# Script to find and disable the local administrator (-500) account
# Used when LAPS is in place for a seperate account

try {

    $account = 'Administrator'  

    $users = Get-LocalUser

    foreach ($user in $users) {
	if ($user.SID.Value.EndsWith("500") -eq $true) {
		$account = $user.name
		}
    }

    $isEnabled = (Get-LocalUser $account -ErrorAction Stop).enabled
}

catch {
    "No such account exists"
}


if ($isEnabled) {  
    Disable-LocalUser $account   
}
