try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$data = ConvertFrom-Csv @'
Item,Quantity,Price,Total Cost
Footballs,9,21.95,197.55
Cones,36,7.99,287.64
Shin Guards,14,10.95,153.3
Turf Shoes,22,79.95,1758.9
Baseballs,68,7.99,543.32
Baseball Gloves,31,65.00,2015.00
Baseball Bats,38,159.00,6042.00
'@

$f = "$env:TEMP\styles.xlsx"
Remove-Item $f -ErrorAction SilentlyContinue

$pkg = $data | Export-Excel -Path $f -AutoSize -PassThru

$ws = $pkg.Workbook.Worksheets["Sheet1"]

Set-ExcelRange -Worksheet $ws -Range "A2:C6" -BackgroundColor PeachPuff -FontColor Purple        -FontSize 12 -Width 12
Set-ExcelRange -Worksheet $ws -Range "D2:D6" -BackgroundColor WhiteSmoke -FontColor Orange -Bold -FontSize 12 -Width 12
Set-ExcelRange -Worksheet $ws -Range "A1:D1" -BackgroundColor BlueViolet -FontColor Wheat        -FontSize 12 -Width 12
Set-ExcelRange -Worksheet $ws -Range "A:A"                                                                    -Width 15

Close-ExcelPackage -ExcelPackage $pkg -Show