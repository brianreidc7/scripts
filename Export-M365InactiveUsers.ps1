<#
    .SYNOPSIS
    Export-M365InactiveUsers.ps1

    .DESCRIPTION
    Export Microsoft 365/Microsoft Entra ID inactive users report.

    .LINK
    www.alitajran.com/export-microsoft-365-inactive-users/

    .NOTES
    Written by: ALI TAJRAN and BRIAN REID (edits)
    Website:    www.alitajran.com
    LinkedIn:   linkedin.com/in/alitajran

    .CHANGELOG
    V1.00, 12/13/2023 - Initial version
    V1.10, 02/26/2024 - Added IsLicensed column
    V1.20, 04/02/2024 - Calculate DaysSinceLastSignIn column from lastSuccessfulSignInDateTime property
    V2.00, 30/10/2024 - Brian Reid - fixed script becuase not working. lastSuccessfulSignInDateTime is an AdditionalProperty and case sensitive
#>

# Export path for CSV file
$CSVPath = ".\InactiveUsers.csv"

# Initialize a List to store the data
$Report = [System.Collections.Generic.List[Object]]::new()

# Connect to Microsoft Graph API
Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All" -NoWelcome

# Get properties
$Properties = @(
    'Id',
    'DisplayName',
    'Mail',
    'UserPrincipalName',
    'UserType',
    'AccountEnabled',
    'SignInActivity',
    'CreatedDateTime',
    'AssignedLicenses',
    'City',
    'Country',
    'UsageLocation'
)

# Get all users along with the properties
$AllUsers = Get-MgUser -All -Property $Properties | Select-Object $Properties

foreach ($User in $AllUsers) {
    # Default values for users who have never signed in
    $LastSuccessfulSignInDate = "Never Signed-in."
    $DaysSinceLastSignIn = "N/A"

    # Check if the user has a successful sign-in date
    if ($User.SignInActivity.AdditionalProperties.lastSuccessfulSignInDateTime) {
        $LastSuccessfulSignInDate = $User.SignInActivity.AdditionalProperties.lastSuccessfulSignInDateTime
        $DaysSinceLastSignIn = (New-TimeSpan -Start $User.SignInActivity.AdditionalProperties.lastSuccessfulSignInDateTime -End (Get-Date)).Days
    }

    # Check if the user is licensed
    $IsLicensed = if ($User.AssignedLicenses) { "Yes" } else { "No" }

    # Collect data in a custom object
    $ReportLine = [PSCustomObject]@{
        Id                       = $User.Id
        UserPrincipalName        = $User.UserPrincipalName
        DisplayName              = $User.DisplayName
        Email                    = $User.Mail
        UserType                 = $User.UserType
        AccountEnabled           = $User.AccountEnabled
        LastSuccessfulSignInDate = $LastSuccessfulSignInDate
        DaysSinceLastSignIn      = $DaysSinceLastSignIn
        CreatedDateTime          = $User.CreatedDateTime
        IsLicensed               = $IsLicensed
	City			 = $User.City
	Country			 = $User.Country
	UsageLocation		 = $User.UsageLocation
    }

    # Add the report line to the List
    $Report.Add($ReportLine)
}

# Display data using Out-GridView
$Report | Out-GridView -Title "Inactive Users"

# Export data to CSV file
try {
    $Report | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8
    Write-Host "Script completed. Results exported to $CSVPath." -ForegroundColor Cyan
}
catch {
    Write-Host "Error occurred while exporting to CSV: $_" -ForegroundColor Red
}