try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$xlSourcefile = "$env:TEMP\Source.xlsx"

Remove-Item -Path $xlSourcefile -ErrorAction Ignore

#Put some simple data in a worksheet and Get an excel package object to represent the file
$excel = 1..10 | Export-Excel $xlSourcefile -PassThru
#Add a new worksheet named 'NewSheet' and copying the sheet that was just made (Sheet1) to the new sheet
Add-WorkSheet -ExcelPackage $excel -WorkSheetname "NewSheet" -CopySource $excel.Workbook.Worksheets["Sheet1"]
#Save and open in Excel
Close-ExcelPackage -ExcelPackage $excel -Show
