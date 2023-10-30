param(
    $PesterTestsPath = "$PSScriptRoot\..\..\..\__tests__\"
)

$xlfile = "$env:Temp\testResults.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$xlparams = @{
    Path              = $xlfile
    InputObject       = (Invoke-Pester -Script $PesterTestsPath -PassThru).TestResult | Sort-Object describe
    WorksheetName     = 'FullResults'

    IncludePivotTable = $true
    PivotRows         = 'Describe'
    PivotColumns      = 'Passed'
    PivotData         = @{'Passed' = 'Count' }

    IncludePivotChart = $true
    ChartType         = 'BarClustered'

    AutoSize          = $true
    AutoFilter        = $true
    Activate          = $true
}

Export-Excel -Show @xlparams