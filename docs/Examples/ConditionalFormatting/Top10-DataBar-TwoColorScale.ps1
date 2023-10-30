try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

$data = ConvertFrom-csv @"
Store,January,February,March,April,May,June
store27,99511,64582,45216,48690,64921,54066
store82,22275,23708,28223,26699,41388,31648
store41,24683,22583,97947,31999,39092,41201
store16,16568,48040,68589,20394,63202,26197
store21,99353,23470,28398,21788,94101,88608
store86,66662,83321,27489,92627,54084,24278
store07,92692,53300,29284,39643,33556,53885
store58,68875,83705,66635,81025,30207,75570
store01,21292,82341,81339,12505,29516,41634
store82,74047,93325,25002,40113,76278,45707
"@

Export-Excel -InputObject $data -Path $xlSourcefile -TableName RawData -WorksheetName RawData
Export-Excel -InputObject $data -Path $xlSourcefile -TableName TopData -WorksheetName StoresTop10Sales 
Export-Excel -InputObject $data -Path $xlSourcefile -TableName Databar -WorksheetName StoresSalesDataBar
Export-Excel -InputObject $data -Path $xlSourcefile -TableName TwoColorScale -WorksheetName StoresSalesTwoColorScale

$xl = Open-ExcelPackage -Path $xlSourcefile

Set-ExcelRange -Worksheet $xl.StoresTop10Sales -Range $xl.StoresTop10Sales.dimension.address -NumberFormat 'Currency' -AutoSize
Set-ExcelRange -Worksheet $xl.StoresSalesDataBar -Range $xl.StoresSalesDataBar.dimension.address -NumberFormat 'Currency' -AutoSize
Set-ExcelRange -Worksheet $xl.StoresSalesTwoColorScale -Range $xl.StoresSalesDataBar.dimension.address -NumberFormat 'Currency' -AutoSize

Add-ConditionalFormatting -Worksheet $xl.StoresTop10Sales -Address $xl.StoresTop10Sales.dimension.address -RuleType Top -ForegroundColor white -BackgroundColor green -ConditionValue 10
Add-ConditionalFormatting -Worksheet $xl.StoresSalesDataBar -Address $xl.StoresSalesDataBar.dimension.address -DataBarColor Red
Add-ConditionalFormatting -Worksheet $xl.StoresSalesTwoColorScale -Address $xl.StoresSalesDataBar.dimension.address -RuleType TwoColorScale

Close-ExcelPackage $xl -Show