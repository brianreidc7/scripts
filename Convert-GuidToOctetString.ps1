param([string]$valuetoconvert) 

function isGUID ($data) { 
    try { 
        $valuetoconvert = [GUID]$data 
        return 1 
    } catch { 
        #$notguid = 1 
        return 0 
    } 
} 
 
function displayhelp  { 
    write-host "Please Supply the value you want converted" 
    write-host "Examples:" 
    write-host "To convert a GUID to an Octet String: Convert-GuidToOctetString.ps1 '234ed3af-58e3-46b0-8cc1-2e1e86b3bf67'" 
    write-host "To convert an Octet String to a GUID: Convert-GuidToOctetString.ps1 'AFD34E23E358B0468CC12E1E86B3BF67'" 
    } 
 
if ($valuetoconvert -eq $NULL) { 
    DisplayHelp 
    return 
} 
# convert a string GUID into an octet (hex) string GUID (like you would use to query AD)
if (isGUID($valuetoconvert)) 
{ 


    [Guid]$g = New-Object Guid;
    if([Guid]::TryParse($valuetoconvert, [ref]$g))
    {
        return ([System.String]::Join('', ($g.ToByteArray() | ForEach-Object { $_.ToString('x2') })).ToUpper());
    }
    else
    {
        throw Exception("Input string is not a valid GUID")
    }

} else { 

    if(32 -eq $valuetoconvert.Length)
    {
        [UInt32]$a = [Convert]::ToUInt32(($valuetoconvert.Substring(6, 2) + $valuetoconvert.Substring(4, 2) + $valuetoconvert.Substring(2, 2) + $valuetoconvert.Substring(0, 2)), 16)
        [UInt16]$b = [Convert]::ToUInt16(($valuetoconvert.Substring(10, 2) + $valuetoconvert.Substring(8, 2)), 16)
        [UInt16]$c = [Convert]::ToUInt16(($valuetoconvert.Substring(14, 2) + $valuetoconvert.Substring(12, 2)), 16)

        [Byte]$d = ([Convert]::ToUInt16($valuetoconvert.Substring(16, 2), 16) -as [byte])
        [Byte]$e = ([Convert]::ToUInt16($valuetoconvert.Substring(18, 2), 16) -as [byte])
        [Byte]$f = ([Convert]::ToUInt16($valuetoconvert.Substring(20, 2), 16) -as [byte])
        [Byte]$g = ([Convert]::ToUInt16($valuetoconvert.Substring(22, 2), 16) -as [byte])
        [Byte]$h = ([Convert]::ToUInt16($valuetoconvert.Substring(24, 2), 16) -as [byte])
        [Byte]$i = ([Convert]::ToUInt16($valuetoconvert.Substring(26, 2), 16) -as [byte])
        [Byte]$j = ([Convert]::ToUInt16($valuetoconvert.Substring(28, 2), 16) -as [byte])
        [Byte]$k = ([Convert]::ToUInt16($valuetoconvert.Substring(30, 2), 16) -as [byte])

        [Guid]$g = New-Object Guid($a, $b, $c, $d, $e, $f, $g, $h, $i, $j, $k)

	Write-Host "Here"
        Write-Host $g.Guid;
    }
    else
    {
        throw Exception("Input string is not a valid octet string GUID")
    }
} else { 
    Write-Host "Value provided not in GUID or Octet String format." 
    DisplayHelp 
} 

