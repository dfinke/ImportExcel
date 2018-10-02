try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$xlFile = "$env:TEMP\mw.xlsx"

Remove-Item -Path $xlFile -ErrorAction Ignore

$leftCsv = @"
MyProp1,MyProp2,Length
a,b,10
c,d,20
"@ | ConvertFrom-Csv

$rightCsv = @"
MyProp1,MyProp2,Length
a,b,10
c,d,21
"@ | ConvertFrom-Csv

Merge-Worksheet -OutputFile $xlFile -ReferenceObject $leftCsv -DifferenceObject $rightCsv -Key Length -Show