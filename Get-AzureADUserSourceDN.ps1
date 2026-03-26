# Determine the source DN of a series of UPN values in a CSV file. CSV must contain one column titled SearchString for this to work

$CSVFileToImport = "IsDuplicate.csv"

$collection = @()

Import-Csv $CSVFileToImport | ForEach-Object{
    $item = New-Object PSObject
    $item | Add-Member -MemberType NoteProperty -Name "SearchString" -Value $_.SearchString
    $item | Add-Member -MemberType NoteProperty -Name "onPremisesDistinguishedName" -Value (get-azureaduser -searchstring $_.SearchString).extensionproperty.onPremisesDistinguishedName
    
    $collection += $item
    $counter++
    Write-Host "Processed" $_.SearchString "(" $counter ")"
    }
$collection | Export-CSV -Path UserSourceDN.csv -NoClobber -NoTypeInformation
Write-Host "Finished, results are in UserSourceDN.csv"
