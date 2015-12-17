$Symbol = "MSFT"

. .\Get-StockInfo.ps1

rm *.xlsx

$chart = New-ExcelChart -XRange Date -YRange Volume `
    -ChartType ColumnStacked `
    -Column 9 -Title "$Symbol Volume"

Get-StockInfo $Symbol 11/2 11/30 | 
    Export-Excel .\stocks.xlsx -Show `
    -AutoSize -AutoNameRange `
    -ExcelChartDefinition $chart