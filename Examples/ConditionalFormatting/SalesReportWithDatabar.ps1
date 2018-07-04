try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

Remove-Item -Path .\test.xlsx -ErrorAction Ignore

$excel = $(
    New-PSItem Jan 1277 (echo Month Sales)
    New-PSItem Feb 1003
    New-PSItem Mar 1105
    New-PSItem Apr 952
    New-PSItem May 770
    New-PSItem Jun 621
) | Export-Excel .\test.xlsx -PassThru -AutoNameRange

$sheet = $excel.Workbook.Worksheets["Sheet1"]

Add-ConditionalFormatting -WorkSheet $sheet -Range "B1:B7" -DataBarColor LawnGreen

$targetRow = $sheet.dimension.Rows + 1

Set-Format -Address $sheet.Cells["A$($targetRow)"] -Value "Total"
Set-Format -Address $sheet.Cells["B$($targetRow)"] -Formula "=Sum(Sales)"

Close-ExcelPackage $excel -Show