try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$xlfile = "$env:Temp\testData.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$r = Get-ChildItem C:\WINDOWS\system32

$BySize=@{}
$r | ForEach-Object{ $BySize.($_.extension)+=$_.length }

$top10BySize = $BySize.GetEnumerator() |
    ForEach-Object{ [PSCustomObject]@{Name=$_.key;Size=[double]$_.value} } |
    Sort-Object size -Descending |
    Select-Object -First 10

$top10ByCount = $r.extension |
    Group-Object |
    Sort-Object count -Descending |
    Select-Object -First 10 Name, count

$top10ByFileSize = $r |
    Sort-Object length -Descending |
    Select-Object -First 10 Name, @{n="Size";e={$_.Length}} #,Extension,Path


$xlPkg = $top10BySize     | Export-Excel -path $xlfile        -WorkSheetname FileInfo              -TableName ExtSize  -PassThru
$xlPkg = $top10ByCount    | Export-Excel -ExcelPackage $xlPkg -WorkSheetname FileInfo -StartRow 13 -TableName ExtCount -PassThru
$xlPkg = $top10ByFileSize | Export-Excel -ExcelPackage $xlPkg -WorkSheetname FileInfo -StartRow 25 -TableName FileSize -PassThru -AutoSize

#worksheets.tables["Name1","Name2"] returns 2 tables. Set-ExcelRange can process those and will set the number format over both
Set-ExcelRange -Range $xlpkg.Workbook.Worksheets[1].Tables["ExtSize","FileSize"] -NumberFormat '0,,"MB"'

$ps = Get-Process | Where-Object Company

$ps |
    Sort-Object handles -Descending |
    Select-Object -First 10 company, handles |
    Export-Excel -ExcelPackage $xlPkg -WorkSheetname Handles -AutoSize -TableName Handles

$ps |
    Sort-Object PM -Descending |
    Select-Object -First 10 company, PM |
    Export-Excel $xlfile -WorkSheetname Handles -AutoSize -TableName PM -StartRow 13 -Show
