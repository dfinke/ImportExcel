try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

#Put some simple data in a worksheet and Get an excel package object to represent the file
$excel = 1..10 | Export-Excel $xlSourcefile -PassThru
#Add a new worksheet named 'NewSheet' and copying the sheet that was just made (Sheet1) to the new sheet
Add-Worksheet -ExcelPackage $excel -WorkSheetname "NewSheet" -CopySource $excel.Workbook.Worksheets["Sheet1"]
#Save and open in Excel
Close-ExcelPackage -ExcelPackage $excel -Show
