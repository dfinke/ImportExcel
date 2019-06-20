function Import-USPS {
    param(
        $TrackingNumber,
        [Switch]$UseDefaultCredentials

    )

    Import-Html "https://tools.usps.com/go/TrackConfirmAction?qtc_tLabels1=$($TrackingNumber)" 0 -UseDefaultCredentials: $UseDefaultCredentials
}

function Import-UPS {
    param(
        $TrackingNumber,
        [Switch]$UseDefaultCredentials
    )

    Import-Html "https://wwwapps.ups.com/WebTracking/track?track=yes&trackNums=$($TrackingNumber)" 0 -UseDefaultCredentials: $UseDefaultCredentials
}