$GroupsOU = "OU=Cloud Licences,OU=Groups,DN=xm,DN=int1"
$LicenceNames = Import-CSV ".\licenceGroups.csv"

Import-Module ActiveDirectory

ForEach ($LicenceName in $LicenceNames)
    {
        Write-Host "Creating " $LicenceName.groupName
        $Description = "Members of this group get the " + $LicenceName.Service + "/" + $LicenceName.SKU + " " + $LicenceName.SubLicence + " licence assigned to them"
        New-ADGroup -Name $LicenceName.GroupName -GroupCategory Security -Description $Description -GroupScope Universal -Path $GroupsOU 
    }

