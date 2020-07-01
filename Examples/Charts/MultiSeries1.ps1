try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

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

$data | ConvertFrom-Csv | Export-Excel -path $xlSourcefile -AutoSize -TableName Impressions
Export-Excel -path $xlSourcefile -worksheetName chartPage -ExcelChartDefinition $c -show
