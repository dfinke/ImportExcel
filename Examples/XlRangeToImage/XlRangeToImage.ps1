try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}


. .\ConvertExcelToImageFile.ps1

$xlFileName = "C:\Temp\testPNG.xlsx"

Remove-Item C:\Temp\testPNG.xlsx -ErrorAction Ignore

$range = @"
Region,Item,Cost
North,Pear,1
South,Apple,2
East,Grapes,3
West,Berry,4
North,Pear,1
South,Apple,2
East,Grapes,3
West,Berry,4
"@ | ConvertFrom-Csv |
    Export-Excel $xlFileName -ReturnRange `
        -ConditionalText (New-ConditionalText Apple), (New-ConditionalText Berry -ConditionalTextColor White -BackgroundColor Purple)

Convert-XlRangeToImage -Path $xlFileName -workSheetname sheet1 -range $range -Show
