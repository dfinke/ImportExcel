# How to convert abbreviate or shorten long numbers in Excel

Remove-Item .\custom.xlsx -ErrorAction SilentlyContinue

$data = $(
    12000
    1000
    2000
    3000
    2400
    3600
    6000
    13000
    40000
    400000
    1000000    
) 

$excel = $data | Export-Excel .\custom.xlsx -PassThru

Set-ExcelRange -Worksheet $excel.Sheet1 -Range "A:A" -NumberFormat '[>999999]#,,"M";#,"K"'

Close-ExcelPackage $excel -Show