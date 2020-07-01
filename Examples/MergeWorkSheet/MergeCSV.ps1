try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

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

Merge-Worksheet -OutputFile $xlSourcefile -ReferenceObject $leftCsv -DifferenceObject $rightCsv -Key Length -Show