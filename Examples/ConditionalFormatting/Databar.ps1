Remove-Item -Path .\test.xlsx -ErrorAction Ignore

$excel = Get-Process |
    Select-Object -Property Name,Company,Handles,CPU,PM,NPM,WS |
    Export-Excel -Path .\test.xlsx -ClearSheet -WorkSheetname "Processes" -PassThru

$sheet = $excel.Workbook.Worksheets["Processes"]
$sheet.Column(1) | Set-Format -Bold -AutoFit
$sheet.Column(2) | Set-Format -Width 29 -WrapText
$sheet.Column(3) | Set-Format -HorizontalAlignment Right -NFormat "#,###"

Set-Format -Address $sheet.Cells["E1:H1048576"]  -HorizontalAlignment Right -NFormat "#,###"
Set-Format -Address $sheet.Column(4)  -HorizontalAlignment Right -NFormat "#,##0.0" -Bold
Set-Format -Address $sheet.Row(1) -Bold -HorizontalAlignment Center

Add-ConditionalFormatting -WorkSheet $sheet -Range "D2:D1048576" -DataBarColor Red
Add-ConditionalFormatting -WorkSheet $sheet -Range "G2:G1048576" -RuleType GreaterThan -ConditionValue "104857600"  -ForeGroundColor Red

foreach ($c in 5..9) {Set-Format -Address $sheet.Column($c)  -AutoFit }

Export-Excel -ExcelPackage $excel -WorkSheetname "Processes" -IncludePivotChart -ChartType ColumnClustered -NoLegend -PivotRows company  -PivotData @{'Name'='Count'}  -Show