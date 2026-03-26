$GroupsOU = "OU=Cloud Licences,OU=Groups,DC=ent,DC=sys,DC=element,DC=com"
# $GroupsOU = "OU=_AAD Groups,DC=bcamericas,DC=com"           # BCAmericas Domain
# $GroupsOU = "OU=_AAD Groups,OU=Exova,DC=exova-eu,DC=local"  # Exova-EU Domain
$LicenceNames = Import-CSV ".\licenceGroups.csv"

Import-Module ActiveDirectory

ForEach ($LicenceName in $LicenceNames)
    {
        Write-Host "Creating " $LicenceName.groupName
        $Description = "Members of this group get the " + $LicenceName.Service + "/" + $LicenceName.SKU + " " + $LicenceName.SubLicence + " licence assigned to them"
        New-ADGroup -Name $LicenceName.GroupName -GroupCategory Security -Description $Description -GroupScope Universal -Path $GroupsOU 
       # Set-ADGroup -Identity $LicenceName.GroupName -Description $Description
    }

