try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$xlSourcefile = "$env:TEMP\Source.xlsx"

Remove-Item $xlSourcefile -ErrorAction Ignore

$excel = 1..10 | Export-Excel $xlSourcefile -PassThru

Add-WorkSheet -ExcelPackage $excel -WorkSheetname "NewSheet" -CopySource $excel.Workbook.Worksheets["Sheet1"]

Close-ExcelPackage -ExcelPackage $excel -Show
