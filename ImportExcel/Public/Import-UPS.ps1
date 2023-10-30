function Import-UPS {
    [CmdletBinding()]
    param(
        $TrackingNumber,
        [Switch]$UseDefaultCredentials
    )

    Import-Html "https://wwwapps.ups.com/WebTracking/track?track=yes&trackNums=$($TrackingNumber)" 0 -UseDefaultCredentials: $UseDefaultCredentials
}