try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

Remove-Item -Path .\test.xlsx -ErrorAction Ignore

$excel = @"
Month,Sales
Jan,1277
Feb,1003
Mar,1105
Apr,952
May,770
Jun,621
"@ | ConvertFrom-csv |
     Export-Excel -Path .\test.xlsx -WorkSheetname Sheet1 -AutoNameRange -PassThru

$sheet = $excel.Workbook.Worksheets["Sheet1"]
Add-ConditionalFormatting -Worksheet $sheet -Range "B1:B7" -DataBarColor LawnGreen

Set-ExcelRange -Address $sheet.Cells["A8"] -Value "Total"
Set-ExcelRange -Address $sheet.Cells["B8"] -Formula "=Sum(Sales)"

Close-ExcelPackage $excel -Show