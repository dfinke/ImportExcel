<#
    Copy a range from WorksheetA to WorksheetB
#>

$data = ConvertFrom-Csv @"
Region,State,Units,Price
West,Texas,927,923.71
North,Tennessee,466,770.67
East,Florida,520,458.68
East,Maine,828,661.24
West,Virginia,465,053.58
North,Missouri,436,235.67
South,Kansas,214,992.47
North,North Dakota,789,640.72
South,Delaware,712,508.55
"@

$xlfile = "./test.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$data | Export-Excel $xlfile -WorksheetName WorksheetA
$data | Export-Excel $xlfile -WorksheetName WorksheetB 

$excel = Open-ExcelPackage $xlfile

# Copy a range from WorksheetA to WorksheetB 
$excel.WorksheetA.Cells["A3:B5"].Copy($excel.WorksheetB.Cells["G3"])

Close-ExcelPackage $excel -Show