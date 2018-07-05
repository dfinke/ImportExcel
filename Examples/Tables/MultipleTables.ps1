try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$xlfile = "testData.xlsx"
Remove-Item *.xlsx

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


$top10BySize     | Export-Excel $xlfile -WorkSheetname FileInfo -TableName ExtSize
$top10ByCount    | Export-Excel $xlfile -WorkSheetname FileInfo -StartRow 13 -TableName ExtCount
$top10ByFileSize | Export-Excel $xlfile -WorkSheetname FileInfo -StartRow 25 -AutoSize -TableName FileSize

$ps = Get-Process | Where-Object Company

$ps |
    Sort-Object handles -Descending |
    Select-Object -First 10 company, handles |
    Export-Excel $xlfile -WorkSheetname Handles -AutoSize -TableName Handles

$ps |
    Sort-Object PM -Descending |
    Select-Object -First 10 company, PM |
    Export-Excel $xlfile -WorkSheetname Handles -AutoSize -TableName PM -StartRow 13 -Show
