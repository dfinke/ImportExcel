#requires -modules "Get-IndexedItem"
[CmdletBinding()]
Param()
Remove-Item ~\documents\music.xlsx -ErrorAction SilentlyContinue
[System.Diagnostics.Stopwatch]$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

#Query system index for .MP3 files in C:\Users, where album artist is non-blank. Leave sorted table with columns of interest in $Music.

Get-IndexedItem "itemtype='.mp3'","AlbumArtist like '%'" -Recurse C:\Users -OutputVariable Music  `
                     -OrderBy  AlbumArtist, AlbumTitle, TrackNumber, Title -NoFiles `
                     -Property AlbumArtist, AlbumTitle, TrackNumber, Title, Duration, Size, SampleRate
Write-Verbose -Message ("Fetched " + $music.Rows.Count + " rows from index: " + $stopwatch.Elapsed.TotalSeconds)
#Send Table in $Music to Excel, format as a table, point $ws to the Worksheet
$excel = Send-SQLDataToExcel -Path ~\documents\music.xlsx -DataTable $music -WorkSheetname Music -TableName Music -Passthru
Write-Verbose -Message ("Inserted into Excel: " + $stopwatch.Elapsed.TotalSeconds)
$ws    = $excel.Music

#Strip "SYSTEM.", "SYSTEM.AUDIO", "SYSTEM.MEDIA", "SYSTEM.MUSIC" from the column headings
#Convert Duration (column 5) from 100ns ticks to days and format as minutes, seconds, decimal
#Format filesize and sample rate nicely
#Autofit the columns.
Set-ExcelRow    -Worksheet $ws -Row 1 -Value {($worksheet.cells[$row,$column].value -replace '^SYSTEM\.','') -replace '^MEDIA\.|^AUDIO\.|^MUSIC\.','' }
Set-ExcelColumn -Worksheet $ws -Column 5 -NumberFormat 'mm:ss.0' -StartRow 2 -Value {$worksheet.cells[$row,$column].value / 864000000000 }
Write-Verbose -Message ("Cells Reset: " + $stopwatch.Elapsed.TotalSeconds)
Set-ExcelColumn -Worksheet $ws -Column 6 -NumberFormat '#.#,,"MB"'
Set-ExcelColumn -Worksheet $ws -Column 7 -NumberFormat '0.0,"KHz"'
$ws.Cells[$ws.Dimension].AutoFitColumns()

#Make a Pivot table for sum of space and count of tracks by artist. Sort by artist, apply formatting to space, give it nice titles.
$pt = Add-PivotTable -PassThru -PivotTableName SpaceUsedByMusic -ExcelPackage $excel -SourceWorkSheet $ws `
                     -PivotRows ALBUMARTIST -PivotData ([ordered]@{"Size"="Sum"; "Duration"="Count"})  -PivotDataToColumn

$pt.RowFields[0].Sort = [OfficeOpenXml.Table.PivotTable.eSortType]::Ascending
$pt.DataFields[0].Format = '#.0,,"MB"'
$pt.DataFields[0].Name   = 'Space Used'
$pt.DataFields[1].Name   = 'Tracks'

#Save the file, and load it into Excel
$stopwatch.Stop()
Write-Verbose -Message ("Pivot Done: " + $stopwatch.Elapsed.TotalSeconds)
Close-ExcelPackage -show $excel
