try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

$data = ConvertFrom-Csv @"
Fruit,Amount
Apples,50
Oranges,20
Bananas,60
Lemons,40
"@

$xl = Export-Excel -InputObject $data -Path $xlSourcefile -PassThru -AutoSize

Set-ExcelRange -Worksheet $xl.Sheet1 -Range D2 -BackgroundColor LightBlue -Value Apples

$rows = $xl.Sheet1.Dimension.Rows
Set-ExcelRange -Worksheet $xl.Sheet1 -Range E2 -Formula "=VLookup(D2,A2:B$($rows),2,FALSE)"

Close-ExcelPackage $xl -Show