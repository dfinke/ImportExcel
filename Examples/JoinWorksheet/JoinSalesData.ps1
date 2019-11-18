try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$xlfile = "$env:temp\AllSales.xlsx"

Remove-Item $xlfile -ErrorAction Ignore

$params = @{
    AutoSize             = $true
    AutoFilter           = $true
    AutoNameRange        = $true
    ExcelChartDefinition = New-ExcelChartDefinition -XRange Item -YRange UnitSold -Title 'Units Sold'
    Path                 = $xlfile
}
#Import 4 sets of sales data from 4 CSV files, using the parameters above.
Import-Csv $PSScriptRoot\NorthSales.csv | Export-Excel -WorkSheetname North @params
Import-Csv $PSScriptRoot\EastSales.csv  | Export-Excel -WorkSheetname East  @params
Import-Csv $PSScriptRoot\SouthSales.csv | Export-Excel -WorkSheetname South @params
Import-Csv $PSScriptRoot\WestSales.csv  | Export-Excel -WorkSheetname West  @params

#Join the 4 worksheets together on a sheet named Allsales, use the same parameters, except for AutoNameRange and ExcelChartDefinition.
$params.Remove("AutoNameRange")
$params.Remove("ExcelChartDefinition")
Join-Worksheet -WorkSheetName AllSales -Show @params