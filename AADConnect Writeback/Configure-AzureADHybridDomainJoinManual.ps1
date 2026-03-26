# See https://docs.microsoft.com/en-us/azure/active-directory/devices/hybrid-azuread-join-manual


###########################################################
# BCAmericas

# Retrieve current settings

$scp = New-Object System.DirectoryServices.DirectoryEntry;

$scp.Path = "LDAP://CN=62a0ff2e-97b9-4513-943f-0d221bd30080,CN=Device Registration Configuration,CN=Services,CN=Configuration,DC=bcamericas,DC=com";

$scp.Keywords;

# Set settings if needed

$verifiedDomain = "element.com"    # Replace this with any of your verified domain names in Azure AD
$tenantID = "97bd1b1c-6ae3-457c-9f9c-a7fc0bc5dfac"    # Replace this with you tenant ID
$configNC = "CN=Configuration,DC=bcamericas,DC=com"    # Replace this with your Active Directory configuration naming context

$de = New-Object System.DirectoryServices.DirectoryEntry
$de.Path = "LDAP://CN=Services," + $configNC
$deDRC = $de.Children.Add("CN=Device Registration Configuration", "container")
$deDRC.CommitChanges()

$deSCP = $deDRC.Children.Add("CN=62a0ff2e-97b9-4513-943f-0d221bd30080", "serviceConnectionPoint")
$deSCP.Properties["keywords"].Add("azureADName:" + $verifiedDomain)
$deSCP.Properties["keywords"].Add("azureADId:" + $tenantID)

$deSCP.CommitChanges()

###########################################################
# Exova-EU

# Retrieve current settings

$scp = New-Object System.DirectoryServices.DirectoryEntry;

$scp.Path = "LDAP://CN=62a0ff2e-97b9-4513-943f-0d221bd30080,CN=Device Registration Configuration,CN=Services,CN=Configuration,DC=exova-eu,DC=local";

$scp.Keywords;

# Set settings if needed

$verifiedDomain = "element.com"    # Replace this with any of your verified domain names in Azure AD
$tenantID = "97bd1b1c-6ae3-457c-9f9c-a7fc0bc5dfac"    # Replace this with you tenant ID
$configNC = "CN=Configuration,DC=exova-eu,DC=local"    # Replace this with your Active Directory configuration naming context

$de = New-Object System.DirectoryServices.DirectoryEntry
$de.Path = "LDAP://CN=Services," + $configNC
$deDRC = $de.Children.Add("CN=Device Registration Configuration", "container")
$deDRC.CommitChanges()

$deSCP = $deDRC.Children.Add("CN=62a0ff2e-97b9-4513-943f-0d221bd30080", "serviceConnectionPoint")
$deSCP.Properties["keywords"].Add("azureADName:" + $verifiedDomain)
$deSCP.Properties["keywords"].Add("azureADId:" + $tenantID)

$deSCP.CommitChanges()


###########################################################
# Bodycote.local (not done - not Enterprise Admin of root domain)

# Retrieve current settings

$scp = New-Object System.DirectoryServices.DirectoryEntry;

$scp.Path = "LDAP://CN=62a0ff2e-97b9-4513-943f-0d221bd30080,CN=Device Registration Configuration,CN=Services,CN=Configuration,DC=bodycote,DC=local";

$scp.Keywords;

# Set settings if needed

$verifiedDomain = "element.com"    # Replace this with any of your verified domain names in Azure AD
$tenantID = "97bd1b1c-6ae3-457c-9f9c-a7fc0bc5dfac"    # Replace this with you tenant ID
$configNC = "CN=Configuration,DC=bodycote,DC=local"    # Replace this with your Active Directory configuration naming context

$de = New-Object System.DirectoryServices.DirectoryEntry
$de.Path = "LDAP://CN=Services," + $configNC
$deDRC = $de.Children.Add("CN=Device Registration Configuration", "container")
$deDRC.CommitChanges()

$deSCP = $deDRC.Children.Add("CN=62a0ff2e-97b9-4513-943f-0d221bd30080", "serviceConnectionPoint")
$deSCP.Properties["keywords"].Add("azureADName:" + $verifiedDomain)
$deSCP.Properties["keywords"].Add("azureADId:" + $tenantID)

$deSCP.CommitChanges()

