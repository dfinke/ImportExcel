try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

$data = ConvertFrom-Csv @"
Week, TotalVisitors
1,11916
2,11665
3,13901
4,15444
5,21592
6,15057
7,26187
8,20662
9,28935
10,32443
"@

$cd = New-ExcelChartDefinition `
    -XRange Week `
    -YRange TotalVisitors `
    -Title "No. Of Visitors" `
    -ChartType ColumnClustered `
    -NoLegend `
    -ChartTrendLine Linear

$data | Export-Excel $xlSourcefile -Show -AutoNameRange -AutoSize -TableName Visitors -ExcelChartDefinition $cd

