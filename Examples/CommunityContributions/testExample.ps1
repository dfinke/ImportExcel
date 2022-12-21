#create simple Excel file
"" | Export-Excel -Path "H:\Test\test.xlsx"

#open excel file and create handler for that file
$excelHandler = Open-ExcelPackage -Path  "H:\Test\test.xlsx"

#handler is an object which represent file in PowerShell code

#Create column with percentage values
#=RAND() -> Excel formula, which generate random value from 0 to 1. Volatile value which changes every Excel file refreshing
Set-ExcelRange -Worksheet $excelHandler.Workbook.Worksheets[1] -Address "A1:A20" -Formula "=RAND()" -NumberFormat 'Percentage' -BackgroundColor "RED"

#creating chart base on this value
Add-ExcelChart -Worksheet $excelHandler.Workbook.Worksheets[1] -Title "Percentages" -ChartType BarClustered

#close Excel chart
Close-ExcelPackage $excelHandler
