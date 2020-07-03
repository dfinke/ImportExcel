# DSUM
# Adds the numbers in a field (column) of records in a list or database that match conditions that you specify.
try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

$data = ConvertFrom-Csv @"
Color,Date,Sales
Red,1/15/2018,250
Blue,1/15/2018,200
Red,1/16/2018,175
Blue,1/16/2018,325
Red,1/17/2018,150
Blue,1/17/2018,300
"@

$xl = Export-Excel -InputObject $data -Path $xlSourcefile -AutoSize -AutoFilter -TableName SalesInfo -AutoNameRange -PassThru

$databaseAddress = $xl.Sheet1.Dimension.Address
Set-Format -Worksheet $xl.Sheet1 -Range C:C -NumberFormat '$##0'

Set-Format -Worksheet $xl.Sheet1 -Range E1 -Value Color
Set-Format -Worksheet $xl.Sheet1 -Range F1 -Value Date
Set-Format -Worksheet $xl.Sheet1 -Range G1 -Value Sales

Set-Format -Worksheet $xl.Sheet1 -Range E2 -Value Red

Set-Format -Worksheet $xl.Sheet1 -Range E4 -Value Sales
Set-Format -Worksheet $xl.Sheet1 -Range F4 -Formula ('=DSUM({0},"Sales",E1:G2)' -f $databaseAddress) -NumberFormat '$##0'

Close-ExcelPackage $xl -Show