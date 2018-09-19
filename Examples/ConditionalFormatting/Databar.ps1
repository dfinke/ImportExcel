try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$path = "$env:temp\test.xlsx"
Remove-Item -Path $path -ErrorAction Ignore

#Export processes, and get an ExcelPackage object representing the file.
$excel = Get-Process |
    Select-Object -Property Name,Company,Handles,CPU,PM,NPM,WS |
    Export-Excel -Path $path -ClearSheet -WorkSheetname "Processes" -PassThru

$sheet = $excel.Workbook.Worksheets["Processes"]

#Apply fixed formatting to columns. Set-Format is an Alias for Set-Excel Range, -NFormat is an alias for numberformat
$sheet.Column(1) | Set-ExcelRange -Bold -AutoFit
$sheet.Column(2) | Set-Format -Width 29 -WrapText
$sheet.Column(3) | Set-Format -HorizontalAlignment Right -NFormat "#,###"

Set-ExcelRange -Range -Address $sheet.Cells["E1:H1048576"]  -HorizontalAlignment Right -NFormat "#,###"
#Set-Format is an alias for Set-ExcelRange
Set-Format -Range   $sheet.Column(4)  -HorizontalAlignment Right -NFormat "#,##0.0" -Bold
#In Set-ExcelRange / Set-Format "-Address" is an alias for "-Range"
Set-Format -Address $sheet.Row(1) -Bold -HorizontalAlignment Center

#Create a Red Data-bar for the values in Column D
Add-ConditionalFormatting -WorkSheet $sheet -Address "D2:D1048576" -DataBarColor Red
# Conditional formatting applies to "Addreses" aliases allow either "Range" or "Address" to be used in Set-ExcelRange or Add-Conditional formatting.
Add-ConditionalFormatting -WorkSheet $sheet -Range  "G2:G1048576" -RuleType GreaterThan -ConditionValue "104857600"  -ForeGroundColor Red

foreach ($c in 5..9) {Set-Format -Address $sheet.Column($c)  -AutoFit }

#Create a pivot and save the file.
Export-Excel -ExcelPackage $excel -WorkSheetname "Processes" -IncludePivotChart -ChartType ColumnClustered -NoLegend -PivotRows company  -PivotData @{'Name'='Count'}  -Show