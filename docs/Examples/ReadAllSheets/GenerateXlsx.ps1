param(
    [Parameter(Mandatory)]
    $path
)

$sheet1 = ConvertFrom-Csv @"
Region,Item,TotalSold
West,melon,27
North,avocado,21
West,kiwi,84
East,melon,23
North,kiwi,8
North,nail,29
North,kiwi,46
South,nail,83
East,pear,10
South,avocado,40
"@

$sheet2 = ConvertFrom-Csv @"
Region,Item,TotalSold
West,lemon,24
North,hammer,41
East,nail,87
West,lemon,68
North,screwdriver,9
North,drill,76
West,lime,28
West,pear,78
North,apple,95
South,melon,40
"@

$sheet3 = ConvertFrom-Csv @"
Region,Item,TotalSold
South,drill,100
East,saw,22
North,saw,5
West,orange,78
East,saw,27
North,screwdriver,57
South,hammer,66
East,saw,62
West,nail,98
West,nail,98
"@

Remove-Item $path -ErrorAction SilentlyContinue

$sheet1 | Export-Excel $path -WorksheetName Sheet1
$sheet2 | Export-Excel $path -WorksheetName Sheet2
$sheet3 | Export-Excel $xlfile -WorksheetName Sheet3

$path