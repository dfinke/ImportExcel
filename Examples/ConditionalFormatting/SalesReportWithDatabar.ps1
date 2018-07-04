try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

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
Add-ConditionalFormatting -WorkSheet $sheet -Range "B1:B7" -DataBarColor LawnGreen

Set-Format -Address $sheet.Cells["A8"] -Value "Total"
Set-Format -Address $sheet.Cells["B8"] -Formula "=Sum(Sales)"

Close-ExcelPackage $excel -Show