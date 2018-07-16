function Show-PesterResults {
    $xlfilename=".\test.xlsx"
    rm $xlfilename -ErrorAction Ignore

    $ConditionalText = @()
    $ConditionalText += New-ConditionalText -Range "Result" -Text failed  -BackgroundColor red   -ConditionalTextColor black
    $ConditionalText += New-ConditionalText -Range "Result" -Text passed  -BackgroundColor green -ConditionalTextColor black
    $ConditionalText += New-ConditionalText -Range "Result" -Text pending -BackgroundColor gray  -ConditionalTextColor black
    
    $xlParams = @{
        Path=$xlfilename
        WorkSheetname = 'PesterTests'
        ConditionalText=$ConditionalText 
        PivotRows = 'Description'
        PivotColumns = 'Result'
        PivotData = @{'Result'='Count'} 
        IncludePivotTable  = $true 
        #IncludePivotChart = $true 
        #NoLegend = $true 
        #ShowPercent = $true 
        #ShowCategory = $true 
        AutoSize = $true 
        AutoNameRange = $true
        AutoFilter = $true
        Show  = $true
    }

    $(foreach($result in (Invoke-Pester -PassThru -Show None).TestResult) {

        [PSCustomObject]@{
            Description = $result.Describe
            Name        = $result.Name
            #Time       = $result.Time
            Result      = $result.Result
            Messge      = $result.FailureMessage
            StackTrace  = $result.StackTrace
        }

    }) | Sort Description | Export-Excel @xlParams 
}