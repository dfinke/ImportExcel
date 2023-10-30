try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

#Export processes, and get an ExcelPackage object representing the file.
$excel = Get-Process |
    Select-Object -Property Name,Company,Handles,CPU,PM,NPM,WS |
    Export-Excel -Path $xlSourcefile -ClearSheet -WorkSheetname "Processes" -PassThru

$sheet = $excel.Workbook.Worksheets["Processes"]

#Apply fixed formatting to columns.  -NFormat is an alias for numberformat
$sheet.Column(1) | Set-ExcelRange -Bold -AutoFit
$sheet.Column(2) | Set-ExcelRange -Width 29 -WrapText
$sheet.Column(3) | Set-ExcelRange -HorizontalAlignment Right -NFormat "#,###"

Set-ExcelRange -Range $sheet.Cells["E1:H1048576"]  -HorizontalAlignment Right -NFormat "#,###"

Set-ExcelRange -Range $sheet.Column(4)  -HorizontalAlignment Right -NFormat "#,##0.0" -Bold
#In Set-ExcelRange "-Address" is an alias for "-Range"
Set-ExcelRange -Address $sheet.Row(1) -Bold -HorizontalAlignment Center

#Create a Red Data-bar for the values in Column D
Add-ConditionalFormatting -Worksheet $sheet -Address "D2:D1048576" -DataBarColor Red
# Conditional formatting applies to "Addreses" aliases allow either "Range" or "Address" to be used in Set-ExcelRange or Add-Conditional formatting.
Add-ConditionalFormatting -Worksheet $sheet -Range  "G2:G1048576" -RuleType GreaterThan -ConditionValue "104857600"  -ForeGroundColor Red

foreach ($c in 5..9) {Set-ExcelRange -Address $sheet.Column($c)  -AutoFit }

#Create a pivot and save the file.
Export-Excel -ExcelPackage $excel -WorkSheetname "Processes" -IncludePivotChart -ChartType ColumnClustered -NoLegend -PivotRows company  -PivotData @{'Name'='Count'}  -Show