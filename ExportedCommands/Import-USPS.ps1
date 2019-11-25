function Import-USPS {
    [CmdletBinding()]
    param(
        $TrackingNumber,
        [Switch]$UseDefaultCredentials

    )

    Import-Html "https://tools.usps.com/go/TrackConfirmAction?qtc_tLabels1=$($TrackingNumber)" 0 -UseDefaultCredentials:$UseDefaultCredentials
}
