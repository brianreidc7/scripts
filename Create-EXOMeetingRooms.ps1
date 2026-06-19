[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [string]$Identity
)

# Update conference rooms based on info in a CSV file

# Brian Reid 2019 - 2026 https://c7solutions.com/2019/08/making-your-office-365-meeting-rooms-accessible
 
# Need to connect to Exchange Online before running this script. Uncomment the line below and change the UPN to your admin account if you want to connect here instead of manually connecting before running the script.
#  Connect-ExchangeOnline -UserPrincipalName exoadmin@tenant.onmicrosoft.com

# This script is provided as-is with no warranty. Use at your own risk. It is recommended to test in a lab environment before running in production.

# This script updates or creates Exchange Online room mailboxes based on the information provided in a CSV file. The CSV file should have the following headers as described in the above blog post.

# If Places updates fails straight after creating a new room mailbox, wait a few minutes and try again. It can take a few minutes for the new mailbox to be discoverable in the directory.
# The -Identity parameter can be used to update a single room mailbox instead of all rooms in the CSV file. The value should be the full email address of the room mailbox.
 
if ([string]::IsNullOrEmpty($CsvPath)) {
    $CsvPath = Read-Host "Enter the path to the rooms CSV file"
}
if (-not (Test-Path -Path $CsvPath -PathType Leaf)) {
    Write-Error "File not found: $CsvPath"
    exit 1
}
$rooms = Import-Csv -Path $CsvPath

if (-not [string]::IsNullOrEmpty($Identity)) {
    $rooms = $rooms | Where-Object { $_.Identity -eq $Identity }
    if ($rooms.Count -eq 0) {
        Write-Error "No entry found in CSV with Identity '$Identity'"
        exit 1
    }
}
 
foreach ($room in $rooms) {
 
    $currentRoom = $room.identity
 
    Write-Host "Updating $($room.DisplayName)"
 
    if ([string]::IsNullOrEmpty($room.street)) { $room.street = "" }
    if ([string]::IsNullOrEmpty($room.city)  ) { $room.city = "" }
    if ([string]::IsNullOrEmpty($room.state) ) { $room.state = "" }
    if ([string]::IsNullOrEmpty($room.postalcode)) { $room.postalcode = "" }
    if ([string]::IsNullOrEmpty($room.CountryOrRegion)) { $room.CountryOrRegion = "" }
 
    if ($room.GeoCoordinates -notlike "*;*") { $room.GeoCoordinates = "" }  # Clear if not formatted properly (expected: latitude;longitude or latitude;longitude;altitude)
 
    if ([string]::IsNullOrEmpty($room.phone)) { $room.phone = "" }
 
    $capacityInt = 0
    if ([int]::TryParse($room.capacity, [ref]$capacityInt)) { $room.capacity = $capacityInt }    # needs to be a number, does not work with null like most of the other values here
 
    if ([string]::IsNullOrEmpty($room.building)) { $room.building = "" }
 
    if ([string]::IsNullOrEmpty($room.Label)) { $room.label = "" }
 
    if ([string]::IsNullOrEmpty($room.AudioDeviceName)) { $room.AudioDeviceName = "" }
    if ([string]::IsNullOrEmpty($room.VideoDeviceName) ) { $room.VideoDeviceName  = "" }
    if ([string]::IsNullOrEmpty($room.DisplayDeviceName)) { $room.DisplayDeviceName = "" }
 
    if ([string]::IsNullOrEmpty($room.IsWheelChairAccessible)) { $room.IsWheelChairAccessible = "False" }
 
    if ([string]::IsNullOrEmpty($room.Floor)) { $room.Floor = "" }    # looks to see if contains no number, not null like most of the other values here
    if ([string]::IsNullOrEmpty($room.FloorLabel)) { $room.FloorLabel = "" }
    if ($room.FloorLabel -eq "" -and $room.Floor -ne "") { $room.FloorLabel = $room.Floor } # Copy the Floor number to FloorLabel if this value is blank
    if ([string]::IsNullOrEmpty($room.tags)) { $room.tags = "" }

    if ([string]::IsNullOrEmpty($room.MTREnabled))       { $room.MTREnabled = "False" }
    if ([string]::IsNullOrEmpty($room.ResourceDelegates)) { $room.ResourceDelegates = "" }
    if ([string]::IsNullOrEmpty($room.TimeZone))         { $room.TimeZone = "" }
    if ([string]::IsNullOrEmpty($room.WorkingHoursStartTime)) { $room.WorkingHoursStartTime = "" }
    if ([string]::IsNullOrEmpty($room.WorkingHoursEndTime))   { $room.WorkingHoursEndTime = "" }
 
    $addressParams = @{
        Street = $room.street
        City = $room.city
        State = $room.state
        PostalCode = $room.postalcode
        CountryOrRegion = $room.CountryOrRegion
    }
 
    $geoParams = @{
        GeoCoordinates = $room.GeoCoordinates
    }
 
    $capacityParams = @{
        Capacity = $room.capacity
    }
 
    $contactParams = @{
        Phone = $room.Phone
        Building = $room.building
        Label = $room.label
    }
 
    $deviceParams = @{
        AudioDeviceName = $room.AudioDeviceName 
        VideoDeviceName = $room.VideoDeviceName 
        DisplayDeviceName = $room.DisplayDeviceName 
    }
 
    $accessibilityParams = @{
        IsWheelChairAccessible = ($room.IsWheelChairAccessible -eq 'True') 
    }
 
    $floorParams = @{
        Floor = $room.Floor 
        FloorLabel = $room.FloorLabel 
        Tags = $room.Tags 
    }

    $managementParams = @{
        MTREnabled  = ($room.MTREnabled -eq 'True')
    }

    # Verify the room mailbox exists; create it if not
    $mailboxCreated = $false
    $mailbox = Get-Mailbox -Identity $currentRoom -ErrorAction SilentlyContinue
    if (-not $mailbox) {
        Write-Host "   Room mailbox not found - creating..."

        # Warn if MTR accounts don't follow the recommended 'mtr-' prefix naming convention
        if ($room.MTREnabled -eq 'True') {
            $localPart = ($room.identity -split '@')[0]
            if (-not $localPart.StartsWith('mtr-')) {
                Write-Warning "   NOTE: MTR account '$($room.identity)' does not follow the recommended 'mtr-' prefix naming convention (e.g. mtr-roomname@domain.com)"
            }
        }

        $newMailboxParams = @{
            MicrosoftOnlineServicesID   = $room.identity
            Name                        = $room.DisplayName
            PrimarySmtpAddress          = $room.identity
            Room                        = $true
        }

        if ($room.MTREnabled -eq 'True') {
            # Generate a unique password for the MTR account
            $passwordChars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+'
            $generatedPassword = -join ((1..16) | ForEach-Object { $passwordChars[(Get-Random -Maximum $passwordChars.Length)] })
            $securePassword = ConvertTo-SecureString -String $generatedPassword -AsPlainText -Force
            $alias = ($room.identity -split '@')[0]

            $newMailboxParams = @{
                MicrosoftOnlineServicesID = $room.identity
                Name                      = $room.DisplayName
                Alias                     = $alias
                Room                      = $true
                EnableRoomMailboxAccount  = $true
                RoomMailboxPassword       = $securePassword
            }
        }

        try {
            $mailbox = New-Mailbox @newMailboxParams -ErrorAction Stop
            $mailboxCreated = $true
            Write-Host "   Room mailbox created."
            if ($room.MTREnabled -eq 'True') {
                Write-Host "   MTR account password: $generatedPassword" -ForegroundColor Blue
                Write-Warning "   NOTE: Ensure the account $($room.identity) is enabled for sign-in in Entra ID"
            }
        } catch {
            Write-Warning "   Failed to create mailbox for '$($room.DisplayName)': $($_.Exception.Message). Skipping."
            continue
        }
    }

    # Check if DisplayName has changed on an existing mailbox and update if so
    if ($mailbox -and $mailbox.DisplayName -ne $room.DisplayName) {
        Write-Host "   DisplayName changed: '$($mailbox.DisplayName)' -> '$($room.DisplayName)' - updating..."
        Set-Mailbox -Identity $currentRoom -DisplayName $room.DisplayName -ErrorAction SilentlyContinue
    }

    # Set-Place can fail with PlaceNotFoundInDirectory shortly after mailbox creation; retry once after 60s
    $setPlaceSuccess = $false
    $setPlaceAttempts = 0
    do {
        $setPlaceAttempts++
        try {
            Set-Place $currentRoom @addressParams -ErrorAction Stop
            $setPlaceSuccess = $true
        } catch {
            if ($_.Exception.Message -like '*PlaceNotFoundInDirectory*' -or $_.Exception.Message -like '*NotFound*') {
                if ($setPlaceAttempts -lt 2) {
                    Write-Host "   Place not yet discoverable in directory - waiting 60 seconds before retrying..."
                    Start-Sleep -Seconds 60
                } else {
                    Write-Warning "   Set-Place failed after retry for '$($room.DisplayName)': $($_.Exception.Message)"
                    $setPlaceSuccess = $true  # exit loop
                }
            } else {
                Write-Warning "   Set-Place (address) failed for '$($room.DisplayName)': $($_.Exception.Message)"
                $setPlaceSuccess = $true  # exit loop, non-retryable error
            }
        }
    } while (-not $setPlaceSuccess)

    if ($room.GeoCoordinates -like "*;*") { 
        Write-Host "   Geo..."
        Set-Place $currentRoom @geoParams   # Only set Geo if there is a valid value to set. Errors if try to set anything else
    }
 
    if ($room.capacity -is [int]) {
        Write-Host "   Capacity..."
        Set-Place $currentRoom @capacityParams # only set capacity if the value is a number or it errors
    }
 
    Write-Host "   Contact..."
    Set-Place $currentRoom @contactParams
 
    Write-Host "   Device..."
    Set-Place $currentRoom @deviceParams
 
    Write-Host "   Accessibility..."
    Set-Place $currentRoom @accessibilityParams
 
    Write-Host "   Floor..."
    Set-Place $currentRoom @floorParams

    Write-Host "   Management..."
    Set-Place $currentRoom @managementParams

    # If this is a Teams Meeting Room (MTR), apply required account and calendar settings
    if ($room.MTREnabled -eq 'True') {
        Write-Host "   Configuring as Teams Meeting Room..."

        # MTR accounts must be enabled for sign-in; room mailboxes are disabled by default
        if (-not $mailboxCreated) {
            Write-Warning "   NOTE: Ensure the account $($room.identity) is enabled for sign-in in Entra ID"
        }

        # Apply Teams Room recommended calendar processing settings
        Set-CalendarProcessing -Identity $currentRoom `
            -AutomateProcessing AutoAccept `
            -AddOrganizerToSubject $false `
            -DeleteComments $false `
            -DeleteSubject $false `
            -RemovePrivateProperty $false `
            -ErrorAction SilentlyContinue

        Write-Warning "   NOTE: Ensure a Teams Rooms license is assigned to $($room.identity)"
    }

    # Set calendar configuration (TimeZone, working hours) if provided
    $calParams = @{}
    if (-not [string]::IsNullOrEmpty($room.TimeZone))              { $calParams['WorkingHoursTimeZone'] = $room.TimeZone }
    if (-not [string]::IsNullOrEmpty($room.WorkingHoursStartTime)) { $calParams['WorkingHoursStartTime'] = [TimeSpan]$room.WorkingHoursStartTime }
    if (-not [string]::IsNullOrEmpty($room.WorkingHoursEndTime))   { $calParams['WorkingHoursEndTime']   = [TimeSpan]$room.WorkingHoursEndTime }
    if ($calParams.Count -gt 0) {
        Write-Host "   Calendar config..."
        Set-MailboxCalendarConfiguration -Identity $currentRoom @calParams -ErrorAction SilentlyContinue
    }

    # Set resource delegates if provided (semicolon-separated list in CSV)
    if (-not [string]::IsNullOrEmpty($room.ResourceDelegates)) {
        Write-Host "   Resource delegates..."
        $delegates = $room.ResourceDelegates -split ';'
        Set-CalendarProcessing -Identity $currentRoom -ResourceDelegates $delegates -ErrorAction SilentlyContinue
    }

    # Add room to City/Building distribution group (RoomList). Create it first if it doesn't exist.
    if (-not [string]::IsNullOrEmpty($room.building)) {
        $roomList = Get-DistributionGroup -Identity $room.building -ErrorAction SilentlyContinue
        if (-not $roomList) {
            Write-Host "   Room list '$($room.building)' not found - creating..."
            $listAlias = ($room.city + "-" + $room.building) -replace '\s+', '-' -replace '[^a-zA-Z0-9\-]', ''
            try {
                New-DistributionGroup -Name $room.building -Alias $listAlias -RoomList -ErrorAction Stop | Out-Null
                Write-Host "   Room list created."
            } catch {
                Write-Warning "   Failed to create room list '$($room.building)': $($_.Exception.Message). Skipping group add."
            }
        }
        Add-DistributionGroupMember $room.building -Member $room.identity -ErrorAction SilentlyContinue
    }
}