try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$xlfile = "$env:TEMP\test.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$data = ConvertFrom-Csv @"
Region,Item,TotalSold
North,melon,38
South,screwdriver,21
South,peach,33
South,saw,81
South,kiwi,70
North,orange,59
North,avocado,25
South,lime,48
South,nail,83
North,apple,2
"@

$styleParams = @{
    FontSize = 13
    Bold     = $true
}

$styles = $(
    New-ExcelStyle -BackgroundColor LightBlue -FontSize 14 -Bold -Range "A1:H1" -HorizontalAlignment Center -Merge

    New-ExcelStyle -BackgroundColor LimeGreen -Range "B10" @styleParams
    New-ExcelStyle -BackgroundColor PeachPuff -Range "B5" @styleParams
    New-ExcelStyle -BackgroundColor Orange -Range "B8" @styleParams
    New-ExcelStyle -BackgroundColor Red -Range "B12" @styleParams
)

$reportTitle = "This is a report Title"
$data | Export-Excel $xlfile -Show -AutoSize -AutoFilter -Title $reportTitle -Style $styles