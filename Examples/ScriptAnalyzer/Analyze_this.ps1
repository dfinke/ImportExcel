$xlfilename = $pwd.path -replace "^.*\\(.*?)\\(.*?)$", '$1-$2.xlsx'
$xlpath     = Join-Path -Path $env:TEMP -ChildPath $xlfilename
Remove-Item -Path $xlpath -ErrorAction SilentlyContinue

$ScriptAnalyzerResults = Invoke-ScriptAnalyzer .

$xlPkg = $ScriptAnalyzerResults | Group-Object -Property RuleName -NoElement | Sort-Object -Property Name |
    Select-Object -Property Name,Count |
        Export-Excel -Path $xlpath -WorksheetName Summary  -AutoSize -PassThru

$params = @{
    WorksheetName     = 'FullResults'
    AutoSize          = $true
    AutoFilter        = $true
    IncludePivotTable = $true
    Activate          = $true
    PivotRows         = 'Severity', 'RuleName'
    PivotData         = @{RuleName = 'Count' }
    IncludePivotChart = $true
    ChartType         = 'BarClustered'
    Show              = $true
}

Export-Excel -ExcelPackage $xlpkg  -InputObject $ScriptAnalyzerResults @params