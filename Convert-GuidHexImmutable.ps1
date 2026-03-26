# From https://scripting.up-in-the.cloud/aadc/the-guid-conversion-carousel.html

#[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
#$title = 'GUID Converter'
#$msg   = 'Enter your GUID / Hex String / Base64:'
#$text = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
 
param([string]$text) 


if ($text.Contains('-')) {
    $guid = $text
    $base64 = [system.convert]::ToBase64String(([GUID]$guid).ToByteArray())
    $hexstring = (([GUID]$guid).ToByteArray() | % ToString X2) -join ' '
} elseif ($text.Contains(' ')) {
    $hexstring = $text
    $guid = [GUID]([byte[]] (-split (($hexstring -replace ' ', '') -replace '..', '0x$& ')))
    $base64 = [system.convert]::ToBase64String([byte[]] (-split (($hexstring -replace ' ', '') -replace '..', '0x$& ')))
} else {
    $base64 = $text
    $guid = [GUID]([system.convert]::FromBase64String($base64))
    $hexstring = ([system.convert]::FromBase64String($base64) | % ToString X2) -join ' '
}
 
""
Write-Host "GUID:                 "$guid.ToString()
Write-Host "ImmutableID (Base64): "$base64
Write-Host "HEX String:           "$hexstring
