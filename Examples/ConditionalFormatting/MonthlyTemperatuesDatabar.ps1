try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

$excel = @"
Month,New York City,Austin Texas,Portland Oregon
Jan,39,61,46
Feb,42,65,51
Mar,50,73,56
Apr,62,80,61
May,72,86,67
Jun,80,92,73
Jul,85,95,80
Aug,84,96,80
Sep,76,90,75
Oct,65,82,63
Nov,54,71,52
Dec,44,63,46
"@ | ConvertFrom-csv |
    Export-Excel -Path $xlSourcefile -WorkSheetname Sheet1 -AutoNameRange -AutoSize -Title "Monthly Temperatures" -PassThru

$sheet = $excel.Workbook.Worksheets["Sheet1"]
Add-ConditionalFormatting -Worksheet $sheet -Range "B1:D14" -DataBarColor CornflowerBlue

Close-ExcelPackage $excel -Show