$xlfile = "testData.xlsx"
rm *.xlsx

$r = dir C:\WINDOWS\system32 

$BySize=@{}
$r | ForEach{ $BySize.($_.extension)+=$_.length }

$top10BySize = $BySize.GetEnumerator() | 
    ForEach{ [PSCustomObject]@{Name=$_.key;Size=[double]$_.value} } | 
    Sort size -Descending | 
    Select -First 10

$top10ByCount = $r.extension | 
    Group | 
    Sort count -Descending | 
    Select -First 10 Name, count

$top10ByFileSize = $r | 
    Sort length -Descending | 
    Select -First 10 Name, @{n="Size";e={$_.Length}} #,Extension,Path


$top10BySize     | Export-Excel $xlfile -WorkSheetname FileInfo -TableName ExtSize
$top10ByCount    | Export-Excel $xlfile -WorkSheetname FileInfo -StartRow 13 -TableName ExtCount
$top10ByFileSize | Export-Excel $xlfile -WorkSheetname FileInfo -StartRow 25 -AutoSize -TableName FileSize

$ps = ps | ? Company 

$ps | 
    sort handles -Descending | 
    select -First 10 company, handles | 
    Export-Excel $xlfile -WorkSheetname Handles -AutoSize -TableName Handles

$ps | 
    sort PM -Descending | 
    select -First 10 company, PM | 
    Export-Excel $xlfile -WorkSheetname Handles -AutoSize -TableName PM -StartRow 13 -Show
