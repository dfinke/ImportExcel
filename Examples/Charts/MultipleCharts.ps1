rm *.xlsx 

$data = @"
ID,Product,Quantity,Price,Total
12001,Nails,37,3.99,147.63
12002,Hammer,5,12.10,60.5
12003,Saw,12,15.37,184.44
12010,Drill,20,8,160  
12011,Crowbar,7,23.48,164.36
"@ | ConvertFrom-Csv 

$xRange = "Product"

$yRange="Price";    $c1 = New-ExcelChart -YRange $yRange -XRange $xRange -Title $yRange -Height 225
$yRange="Total";    $c2 = New-ExcelChart -YRange $yRange -XRange $xRange -Title $yRange -Row 9 -Column 15 -Height 225 
$yRange="Quantity"; $c3 = New-ExcelChart -YRange $yRange -XRange $xRange -Title $yRange -Row 15 -Height 225 

$data | 
    Export-Excel -ExcelChartDefinition $c1,$c2,$c3 Tools.xlsx -Show -AutoFilter -AutoNameRange -AutoSize
