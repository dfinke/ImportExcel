$xlfile = "$env:TEMP\test.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$data = ConvertFrom-Csv @"
Fruit,Amount
Apples,50
Oranges,20
Bananas,60
Lemons,40
"@

$xl = Export-Excel -InputObject $data -Path $xlfile -PassThru -AutoSize

Set-ExcelRange -Worksheet $xl.Sheet1 -Range D2 -BackgroundColor LightBlue -Value Apples

$Rows = $xl.Sheet1.Dimension.Rows
Set-ExcelRange -Worksheet $xl.Sheet1 -Range E2 -Formula "=VLookup(D2,A2:B$($Rows),2,FALSE)"

Close-ExcelPackage $xl -Show