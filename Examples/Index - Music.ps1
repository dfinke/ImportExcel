del ~\documents\music.xlsx 
$excel = Get-IndexedItem "itemtype='.mp3'","AlbumArtist like '%'","RatingText <> '1 star'" -NoFiles -orderby AlbumArtist,AlbumTitle,TrackNumber -path c:\users -Recurse -Property AlbumArtist,Duration,title,EncodingBitrate,SampleRate,AlbumTitle,TrackNumber, Size | 
    Select-Object -Property AlbumArtist, AlbumTitle, TrackNumber, Title, Duration, SampleRate, EncodingBitRate, Size  | Export-excel -path ~\documents\music.xlsx -WorksheetName Music -AutoNameRange -AutoSize -BoldTopRow -FreezeTopRow -PassThru  
$ws = $excel.Workbook.Worksheets[1]
Set-Column -Worksheet $ws -Column 6 -NumberFormat '0,"KHz"' 
Set-Column -Worksheet $ws -Column 7 -NumberFormat '0,"Kbits/Sec"' -Width 18
Set-Column -Worksheet $ws -Column 8 -NumberFormat '#.#,,"MB"' -Width 7
$pt = Add-PivotTable -PivotTableName SpaceUsedByMusic -ExcelPackage $excel -SourceWorkSheet $ws -PivotRows ALBUMARTIST -PivotData @{"Size"="Sum"} -PivotNumberFormat '#.#,,"MB"' -Activate -PassThru
$pt.RowFields[0].Sort = [OfficeOpenXml.Table.PivotTable.eSortType]::Ascending

$a = $ws.Dimension.address
Add-ExcelTable -Range $ws.cells[$a] -TableStyle Light1 -TableName Musictable -ShowFilter:$false -ShowTotal -ShowFirstColumn
Add-ConditionalFormatting -Address $ws.Names[1]              -RuleType ContainsText -ConditionValue "Hits" -ForeGroundColor Blue
Add-ConditionalFormatting -Address $ws.Cells["Albumartist"]  -RuleType ContainsText -ConditionValue "Numan" -ForeGroundColor red
Close-ExcelPackage -show $excel