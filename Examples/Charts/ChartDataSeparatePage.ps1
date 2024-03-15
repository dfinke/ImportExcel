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

$xlfile = "$PSScriptRoot\spike.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$xlpkg = $data | Export-Excel $xlfile -WorksheetName Data -AutoNameRange -PassThru

$null = Add-Worksheet -ExcelPackage $xlpkg -WorksheetName Summary -Activate

$params = @{
    Worksheet = $xlpkg.Summary
    Title     = "Sales by Region"
    ChartType = 'ColumnClustered'
    
    # XRange    = "Data!A2:A10"
    # YRange    = "Data!C2:C10"

    XRange    = 'Data!Region'
    YRange    = 'Data!Units'
}

Add-ExcelChart  @params

Close-ExcelPackage $xlpkg -Show
