#requires -modules "getSql"

try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}
#download f1Results from https://1drv.ms/f/s!AhfYu7-CJv4egbt5FD7Cdxi8jSz3aQ  and update the path below
Get-SQL -Session f1 -Excel  -Connection C:\Users\mcp\OneDrive\Public\F1\f1Results.xlsx -showtables -Verbose

Remove-Item .\demo3.xlsx
$session = $DbSessions["f1"]

$SQL     =  "SELECT top 25 DriverName,  Count(RaceDate) as Races ,
                    Count(Win) as Wins, Count(Pole) as Poles, Count(FastestLap) as Fastlaps
             FROM Results GROUP BY DriverName
             order by (count(win)) desc"
$Excel   =  Send-SQLDataToExcel -SQL $sql -Session $session -path .\demo3.xlsx -WorkSheetname "Winners" -AutoSize -AutoNameRange -BoldTopRow -FreezeTopRow -Passthru

$ws = $Excel.Workbook.Worksheets["Winners"]

Set-ExcelRow    -Worksheet $ws -Heading "Average"     -Value {"=Average($columnName`2:$columnName$endrow)"}  -NumberFormat "0.0"  -Bold
Set-ExcelColumn  -Worksheet $ws -Heading "WinsToPoles" -Value {"=D$row/C$row"}           -Column 6            -AutoSize -AutoNameRange
Set-ExcelColumn -Worksheet $ws -Heading "WinsToFast"  -Value {"=E$row/C$row"}           -Column 7            -AutoSize -AutoNameRange

Set-ExcelRange -Worksheet $ws -Range "F2:G50" -NumberFormat "0.0%"
$chart = New-ExcelChartDefinition  -NoLegend -ChartType XYScatter -XRange WinsToFast -YRange WinsToPoles -Column 7 -Width 2000 -Height 700 -Title "Poles vs fastlaps"
Export-Excel -ExcelPackage $Excel -WorkSheetname "Winners" -ExcelChartDefinition $chart -Show