# Creates a worksheet, addes a chart and then a Linear trendline
try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

$data = ConvertFrom-Csv @"
Region,Item,TotalSold
West,screws,60
South,lemon,48
South,apple,71
East,screwdriver,70
East,kiwi,32
West,screwdriver,1
South,melon,21
East,apple,79
South,apple,68
South,avocado,73
"@

$cd = New-ExcelChartDefinition -XRange Region -YRange TotalSold -ChartType ColumnClustered -ChartTrendLine Linear
$data | Export-Excel $xlSourcefile -ExcelChartDefinition $cd -AutoNameRange -Show
