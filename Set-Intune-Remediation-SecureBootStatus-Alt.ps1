<#
.SYNOPSIS
    Remediation script to initiate the 2026 Secure Boot certificate update.
    Includes guardrails to prevent unnecessary triggers on pending-reboot or blocked devices.

Downloaded from https://www.reddit.com/r/Intune/s/6sjjoIEGfK

#>

$ErrorActionPreference = "SilentlyContinue"

$regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing"
$status = (Get-ItemProperty -Path $regPath -Name "UEFICA2023Status" -ErrorAction SilentlyContinue).UEFICA2023Status
$errorCode = (Get-ItemProperty -Path $regPath -Name "UEFICA2023Error" -ErrorAction SilentlyContinue).UEFICA2023Error

# Guardrail 1: Do not touch if pending reboot (2147942750 = 0x8007015E)
if ($status -eq "InProgress" -and $errorCode -eq 2147942750) {
    Write-Output "No action taken. Device is safely pending a user reboot."
    exit 0
}

# Guardrail 2: Do not hammer if firmware is actively blocking it
if ($errorCode -and $errorCode -ne 0 -and $errorCode -ne 2147942750) {
    Write-Output "No action taken. Device requires an OEM BIOS update before remediation can succeed."
    exit 0
}

Write-Output "Initiating Secure Boot certificate deployment..."

try {
    # Set the trigger key to deploy all needed certificates and update the boot manager (0x5944)
    $triggerPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot"
    if (!(Test-Path $triggerPath)) { 
        New-Item -Path $triggerPath -Force | Out-Null 
    }
    Set-ItemProperty -Path $triggerPath -Name "AvailableUpdates" -Value 0x5944 -Type DWord -Force

    # Trigger the native Windows evaluation task
    $taskName = "\Microsoft\Windows\PI\Secure-Boot-Update"
    Start-ScheduledTask -TaskName $taskName -ErrorAction Stop

    Write-Output "Success: Triggered the Secure-Boot-Update task. Will re-evaluate on next sync."
    exit 0

} catch {
    Write-Output "Remediation Failed: Could not set registry keys or trigger task. $($_.Exception.Message)"
    exit 1
}