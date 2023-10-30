$xlfile = "$env:TEMP\MultipleSheets.xlsx"

.\GenerateXlsx.ps1 $xlfile
.\Get-ExcelSheets.ps1 $xlfile