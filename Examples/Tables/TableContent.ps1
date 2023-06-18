try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$table1 = ConvertFrom-Csv @"
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
$table2 = ConvertFrom-Csv @"
Region,State,Units,Price,Index
West,Texas,927,923.71,1
North,Tennessee,466,770.67,1
East,Florida,520,458.68,1
East,Maine,828,661.24,1
West,Virginia,465,053.58,1
North,Missouri,436,235.67,1
South,Kansas,214,992.47,1
North,North Dakota,789,640.72,1
South,Delaware,712,508.55,1
"@

# Get rid of pre-exisiting sheet
$xlsx = "TableContent.xlsx"
if (test-path $xlsx) {remove-item $xlsx}

# Add 3 sheets: "tab1" - 2 tables, "tab1" - 1 table, and "empty"
$excel = $table1 | Export-Excel -path $xlsx -WorksheetName tab1 -tablename table1 -PassThru
$excel = $table2 | Export-Excel -ExcelPackage $excel -WorksheetName tab1 -tablename table2 -StartRow ($table1.count+5) -PassThru
# Get table names from the first worksheet
Get-ExcelTable -ExcelPackage $excel
$excel = $table1 | Export-Excel -ExcelPackage $excel -tablename table21 -WorksheetName tab2 -StartRow 3 -StartColumn 3 -PassThru
$null = Add-Worksheet -ExcelPackage $excel -WorksheetName Empty
Close-ExcelPackage $excel -Show

# Get all table names
Get-ExcelTable -path $xlsx

# Get all table names including empty worksheets
Get-ExcelTable -path $xlsx -IncludeEmptySheet

# Get table contents
Get-ExcelTable -path $xlsx -Content

# Get table contents from selected worksheets
Get-ExcelTable -path $xlsx -Content -WorksheetName tab1

# Get table content from the first table of each worksheet
0 | Get-ExcelTable -path $xlsx -Content
