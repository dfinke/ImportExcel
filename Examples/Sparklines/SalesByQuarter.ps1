try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$xlfile = "$env:TEMP\SalesByQuarter.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$data = ConvertFrom-Csv @"
Region,Q1,Q2,Q3,Q4,YTDPerformance
Asia,1400,7200,5700,6900
Europe,3400,2300,9400,7300
Midwest,4700,9300,3700,8600
Northeast,2300,4300,4600,5600
"@

$excel = $data | Export-Excel $xlfile -Passthru -AutoSize -TableName SalesByQuarter

$ws = $excel.Sheet1

Set-ExcelRange -Worksheet $ws -Range "B2:E5" -NumberFormat "$#,##0" -AutoSize
$sparkLineType = "line"
$null = $ws.SparklineGroups.Add( $sparkLineType, $ws.Cells["F2"], $ws.Cells["B2:E2"] )
$null = $ws.SparklineGroups.Add( $sparkLineType, $ws.Cells["F3"], $ws.Cells["B3:E3"] )
$null = $ws.SparklineGroups.Add( $sparkLineType, $ws.Cells["F4"], $ws.Cells["B4:E4"] )
$null = $ws.SparklineGroups.Add( $sparkLineType, $ws.Cells["F5"], $ws.Cells["B5:E5"] )

Close-ExcelPackage $excel -Show