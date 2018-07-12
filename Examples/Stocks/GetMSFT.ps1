<#
    Revisit I think yahoo deprecated their service
#>

# try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

# $Symbol = "MSFT"

# . .\Get-StockInfo.ps1

# Remove-Item *.xlsx -ErrorAction Ignore

# $chart = New-ExcelChart -XRange Date -YRange Volume `
#     -ChartType ColumnStacked `
#     -Column 9 -Title "$Symbol Volume"

# Get-StockInfo $Symbol 11/2 11/30 |
#     Export-Excel .\stocks.xlsx -Show `
#     -AutoSize -AutoNameRange `
#     -ExcelChartDefinition $chart