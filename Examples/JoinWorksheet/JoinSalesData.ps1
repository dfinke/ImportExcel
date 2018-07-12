try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$xlfile = "$env:temp\AllSales.xlsx"

Remove-Item $xlfile -ErrorAction Ignore

$params = @{
    AutoSize             = $true
    AutoFilter           = $true
    AutoNameRange        = $true
    ExcelChartDefinition = New-ExcelChartDefinition -XRange Item -YRange UnitSold -Title 'Units Sold'
    Path                 = $xlfile
}

Import-Csv $PSScriptRoot\NorthSales.csv | Export-Excel -WorkSheetname North @params
Import-Csv $PSScriptRoot\EastSales.csv  | Export-Excel -WorkSheetname East  @params
Import-Csv $PSScriptRoot\SouthSales.csv | Export-Excel -WorkSheetname South @params
Import-Csv $PSScriptRoot\WestSales.csv  | Export-Excel -WorkSheetname West  @params

$params.Remove("AutoNameRange")
$params.Remove("ExcelChartDefinition")
Join-Worksheet -WorkSheetName AllSales -Show @params