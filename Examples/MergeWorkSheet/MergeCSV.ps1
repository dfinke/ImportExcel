try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$xlFile = "$env:TEMP\mw.xlsx"

Remove-Item $xlFile -ErrorAction Ignore

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