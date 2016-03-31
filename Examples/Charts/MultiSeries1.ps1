rm temp.xlsx -ErrorAction Ignore

$data = @"
A,B,C,Date
2,1,1,2016-03-29
5,10,1,2016-03-29
"@ | ConvertFrom-Csv 

$c = New-ExcelChart -Title Impressions `
    -ChartType Line -Header "Something" `
    -XRange "Impressions[Date]" `
    -YRange @("Impressions[B]","Impressions[A]") `
    -SeriesHeader 'B data','A data'

$data | 
    Export-Excel temp.xlsx -AutoSize -TableName Impressions -Show -ExcelChartDefinition $c