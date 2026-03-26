<#
.SYNOPSIS
    Detection script to evaluate the deployment status of 2026 Secure Boot certificates.
    Provides formatted output for clean Intune reporting.

Downloaded from https://www.reddit.com/r/Intune/s/6sjjoIEGfK

#>

$ErrorActionPreference = "SilentlyContinue"

# Check if Secure Boot is enabled on the OS level
if (!(Confirm-SecureBootUEFI)) {
    Write-Output "Status: [UNSUPPORTED] - Secure Boot is disabled or not supported on this device."
    exit 1 
}

$regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing"
$status = (Get-ItemProperty -Path $regPath -Name "UEFICA2023Status" -ErrorAction SilentlyContinue).UEFICA2023Status
$errorCode = (Get-ItemProperty -Path $regPath -Name "UEFICA2023Error" -ErrorAction SilentlyContinue).UEFICA2023Error
$errorEvent = (Get-ItemProperty -Path $regPath -Name "UEFICA2023ErrorEvent" -ErrorAction SilentlyContinue).UEFICA2023ErrorEvent

# Format the error code into a clean Hex string for the Intune console
$hexError = if ($null -ne $errorCode) { "0x{0:X8}" -f $errorCode } else { "None" }

# 1. Check for the specific "Pending Reboot" state (0x8007015E / 2147942750)
if ($status -eq "InProgress" -and $hexError -eq "0x8007015E") {
    Write-Output "Status: [PENDING REBOOT] - Certs applied. Waiting on user to reboot to swap the Boot Manager."
    exit 1 # Exiting 1 keeps it flagged as an "Issue Found" in Intune until the reboot happens
}

# 2. Check for actual Firmware Errors
if ($errorCode -and $errorCode -ne 0 -and $hexError -ne "0x8007015E") {
    Write-Output "Status: [FIRMWARE BLOCKED] - BIOS rejected the payload. OEM update required. Error: $hexError (Event: $errorEvent)"
    exit 1 
}

# 3. Evaluate standard deployment states
if ($status -eq "Updated") {
    Write-Output "Status: [COMPLIANT] - The 2026 certificates are successfully applied."
    exit 0 # Healthy
} elseif ($status -eq "InProgress") {
    Write-Output "Status: [IN PROGRESS] - The update is actively processing. Error code: $hexError"
    exit 1 
} elseif ($status -eq "NotStarted" -or $null -eq $status) {
    Write-Output "Status: [NOT STARTED] - The update payload has not been initiated."
    exit 1 
} else {
    Write-Output "Status: [UNKNOWN] - Raw Status: $status | Error: $hexError"
    exit 1
}