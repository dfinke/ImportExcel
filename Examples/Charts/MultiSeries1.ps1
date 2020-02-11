try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

Remove-Item temp.xlsx -ErrorAction Ignore

$data = @"
A,B,C,Date
2,1,1,2016-03-29
5,10,1,2016-03-29
"@

$c = New-ExcelChartDefinition -Title Impressions `
    -ChartType Line   `
    -XRange "Impressions[Date]" `
    -YRange @("Impressions[B]","Impressions[A]") `
    -SeriesHeader 'B data','A data' `
    -Row 0 -Column 0

$data | ConvertFrom-Csv |   Export-Excel -path temp.xlsx -AutoSize -TableName Impressions
Export-Excel -path temp.xlsx -worksheetName chartPage -ExcelChartDefinition $c  -show
