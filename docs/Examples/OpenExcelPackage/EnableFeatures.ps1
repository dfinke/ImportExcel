# How to use Enable-ExcelAutoFilter and Enable-ExcelAutofit

try { Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 } catch { throw ; return }

$data = ConvertFrom-Csv @"
RegionInfo,StateInfo,Units,Price
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

$xlfile = "$PSScriptRoot\enableFeatures.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$data | Export-Excel $xlfile

$excel = Open-ExcelPackage $xlfile

Enable-ExcelAutoFilter -Worksheet $excel.Sheet1 
Enable-ExcelAutofit -Worksheet $excel.Sheet1
 
Close-ExcelPackage $excel -Show