# Check if Secure Boot UEFI database contains the 2023 certificate
# From https://scloud.work/intune-secure-boot-certificate-updates/

try {
    $db = Get-SecureBootUEFI -Name db
    $dbString = [System.Text.Encoding]::ASCII.GetString($db.Bytes)
} catch {
    Write-Output "Error: Unable to read Secure Boot UEFI DB. Device may not support Secure Boot or access is blocked."
    exit 1
}

# Match for the new certificate
$match = $dbString -match 'Windows UEFI CA 2023'

if ($match) {
    Write-Output "Compliant: Windows UEFI CA 2023 is present in the Secure Boot database."
    exit 0
} else {
    Write-Output "Non-Compliant: Windows UEFI CA 2023 not found in the Secure Boot database."
    exit 1
}
