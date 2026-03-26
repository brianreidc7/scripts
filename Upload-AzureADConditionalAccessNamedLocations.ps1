# Script from https://goodworkaround.com/2019/11/09/populating-azure-ad-named-and-trusted-locations-using-graph/

$url = Read-Host "Paste Graph Explorer url"
$excelFile = "AzureAD-ConditionalAccess-LocationIP.xlsx"
 
# Extract access token and create header
$accessToken = ($url -split "access_token=" | select -Index 1) -split "&" | select -first 1
$headers = @{"Authorization" = "Bearer $accessToken"}
 
# Get existing named locations
$_namedLocationsAzureAD = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/conditionalAccess/namedLocations/" -Headers $headers
$namedLocationsAzureAD = $_namedLocationsAzureAD.value | foreach{[PSCustomObject]@{id = $_.id; displayName=$_.displayName; isTrusted = $_.isTrusted; ipRanges = @($_.ipranges.cidraddress)}}
 
# Get locations from Excel
$namedLocationsExcel = @{}
Import-Excel -Path $excelFile | ? Location | Foreach {
    $IP = $_.IP 
    if($IP -notlike "*/*" -and $IP -like "*.*") {
        Write-Verbose "Changed $IP to $IP/32" -Verbose
        $IP = $IP + "/32"
    } elseif($IP -notlike "*/*" -and $IP -like "*:*") {
        Write-Verbose "Changed $IP to $IP/128" -Verbose
        $IP = $IP + "/128"
    }
 
    $namedLocationsExcel[$_.Location] += @($IP)
}
 
# Work in each named location in Excel
$namedLocationsExcel.Keys | Foreach {
    Write-Verbose -Message "Working on location $($_) from Excel" -Verbose
 
    $Body = @{
        "@odata.type" = "#microsoft.graph.ipNamedLocation"
        displayName = $_
        isTrusted = $true
        ipRanges = @($namedLocationsExcel[$_] | Foreach {
            if($_ -like "*.*") {
                @{
                    "@odata.type" = "#microsoft.graph.iPv4CidrRange"
                    cidrAddress = $_
                }
            } else {
                @{
                    "@odata.type" = "#microsoft.graph.iPv6CidrRange"
                    cidrAddress = $_
                }
            }
        })
    } | ConvertTo-Json -Depth 4
 
     $Body
 
    $existingLocation = $namedLocationsAzureAD | ? displayName -eq $_
    if($existingLocation) {
        $key = $_
        if(($existingLocation.ipRanges | where{$_ -notin $namedLocationsExcel[$key]}) -or ($namedLocationsExcel[$key] | where{$_ -notin $existingLocation.ipRanges})) {
            Write-Verbose "Location $($_) has wrong subnets -&amp;gt; updating" -Verbose
            Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/conditionalAccess/namedLocations/$($existingLocation.id)" -Headers $headers -Method Patch -Body $Body -ContentType "application/json" | Out-Null
        }
         
    } else {
        Write-Verbose "Location $($_) does not exist -&amp;gt; creating" -Verbose
        Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/conditionalAccess/namedLocations" -Headers $headers -Method Post -Body $Body -ContentType "application/json" | Out-Null
    }
}