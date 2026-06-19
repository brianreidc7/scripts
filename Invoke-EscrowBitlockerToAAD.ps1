<#
.SYNOPSIS
    Escrow (Backup) the existing Bitlocker key protectors to Entra ID (Intune)

.DESCRIPTION
    This script will verify the presence of existing recovery keys and have them escrowed (backed up) to Entra ID
    Great for switching away from MBAM on-prem to using Intune and Azure AD for Bitlocker key management

.INPUTS
    None

.NOTES
    Version       : 1.2
    Author        : Michael Mardahl
    Twitter       : @michael_mardahl
    Blogging on   : www.msendpointmgr.com
    Creation Date : 11 January 2021
    Purpose/Change: Initial script
    License       : MIT (Leave author credits)

.EXAMPLE
    Execute script as system or administrator
    .\Invoke-EscrowBitlockerToAAD.ps1

.NOTES
    If there is a policy mismatch, then you might get errors from the built-in cmdlet BackupToAAD-BitLockerKeyProtector.
    So I have wrapped the cmdlet in a try/catch in order to supress the error. This means that you will have to manually verify that the key was actually escrowed.
    Check MSEndpointMgr.com for solutions to get reporting stats on this.

    v1.1 Brian Reid to support multiple keys and create a new key if needed
    v1.2 Brian Reid - if no recovery keys found, but disk is encrypted, create a new recovery key and escrow that to Entra ID

#>

#region declarations

$DriveLetter = $env:SystemDrive

#endregion declarations

#region functions

function Test-Bitlocker ($BitlockerDrive) {
    #Tests the drive for existing Bitlocker keyprotectors
    try {
        Get-BitLockerVolume -MountPoint $BitlockerDrive -ErrorAction Stop
    } catch {
        Write-Output "Bitlocker was not found protecting the $BitlockerDrive drive. Terminating script!"
        return $false
    }
}

function Get-KeyProtectorId ($BitlockerDrive) {
    #fetches the key protector ID of the drive
    $BitLockerVolume = Get-BitLockerVolume -MountPoint $BitlockerDrive
    $KeyProtector = $BitLockerVolume.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }
    return $KeyProtector.KeyProtectorId
}

function Invoke-BitlockerEscrow ($BitlockerDrive,$BitlockerKey) {
    #Escrow the key into Entra ID

    foreach ($Key in $BitlockerKey) {

        try {
            BackupToAAD-BitLockerKeyProtector -MountPoint $BitlockerDrive -KeyProtectorId $Key -ErrorAction Stop
            Write-Output "Attempted to escrow key in Entra ID - Please verify manually!"
            
        } catch {
            Write-Error "This should never have happened? Debug me!"
            return $false
        }

    }
    return $true
}

#endregion functions

#region execute

if (-not (Test-Bitlocker -BitlockerDrive $DriveLetter)) {
    exit 11
}
$KeyProtectorId = Get-KeyProtectorId -BitlockerDrive $DriveLetter

if ($KeyProtectorId.Count -gt 0) {
    # Recovery key(s) found so back them up
    if (-not (Invoke-BitlockerEscrow -BitlockerDrive $DriveLetter -BitlockerKey $KeyProtectorId)) {
        exit 12
    }
} else {
    # No recovery keys found, so add one and backup this recovery key
    Add-BitLockerKeyProtector -MountPoint $DriveLetter -RecoveryPasswordProtector

    $KeyProtectorId = Get-KeyProtectorId -BitlockerDrive $DriveLetter
    if (-not (Invoke-BitlockerEscrow -BitlockerDrive $DriveLetter -BitlockerKey $KeyProtectorId)) {
        exit 13
    }
}
exit 0


#endregion execute
