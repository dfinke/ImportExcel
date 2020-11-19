function Show-PesterResults {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Justification="No suitable singular")]
    Param()
    $xlfilename = ".\test.xlsx"
    Remove-Item $xlfilename -ErrorAction Ignore

    $ConditionalText = @()
    $ConditionalText += New-ConditionalText -Range "Result" -Text failed  -BackgroundColor red   -ConditionalTextColor black
    $ConditionalText += New-ConditionalText -Range "Result" -Text passed  -BackgroundColor green -ConditionalTextColor black
    $ConditionalText += New-ConditionalText -Range "Result" -Text pending -BackgroundColor gray  -ConditionalTextColor black

    $xlParams = @{
        Path              = $xlfilename
        WorkSheetname     = 'PesterTests'
        ConditionalText   = $ConditionalText
        PivotRows         = 'Result', 'Name'
        PivotData         = @{'Result' = 'Count'}
        IncludePivotTable = $true
        AutoSize          = $true
        AutoNameRange     = $true
        AutoFilter        = $true
        Show              = $true
    }

    $(foreach ($result in (Invoke-Pester -PassThru -Show None).TestResult) {
            [PSCustomObject]@{
                Description = $result.Describe
                Name        = $result.Name
                Result      = $result.Result
                Messge      = $result.FailureMessage
                StackTrace  = $result.StackTrace
            }
        }) | Sort-Object Description | Export-Excel @xlParams
}