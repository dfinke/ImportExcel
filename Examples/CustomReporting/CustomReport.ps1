try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

$data = @"
From,To,RDollars,RPercent,MDollars,MPercent,Revenue,Margin
Atlanta,New York,3602000,.0809,955000,.09,245,65
New York,Washington,4674000,.105,336000,.03,222,16
Chicago,New York,4674000,.0804,1536000,.14,550,43
New York,Philadelphia,12180000,.1427,-716000,-.07,321,-25
New York,San Francisco,3221000,.0629,1088000,.04,436,21
New York,Phoneix,2782000,.0723,467000,.10,674,33
"@ | ConvertFrom-Csv

$data | Export-Excel $xlSourcefile -AutoSize

$excel = Open-ExcelPackage $xlSourcefile

$sheet1 = $excel.Workbook.Worksheets["sheet1"]

$sheet1.View.ShowGridLines = $false
$sheet1.View.ShowHeaders = $false

Set-ExcelRange -Address $sheet1.Cells["C:C"] -NumberFormat "$#,##0" -WrapText -HorizontalAlignment Center
Set-ExcelRange -Address $sheet1.Cells["D:D"] -NumberFormat "#.#0%"  -WrapText -HorizontalAlignment Center

Set-ExcelRange -Address $sheet1.Cells["E:E"] -NumberFormat "$#,##0" -WrapText -HorizontalAlignment Center
Set-ExcelRange -Address $sheet1.Cells["F:F"] -NumberFormat "#.#0%"  -WrapText -HorizontalAlignment Center

Set-ExcelRange -Address $sheet1.Cells["G:H"] -WrapText -HorizontalAlignment Center

## Insert Rows/Columns
$sheet1.InsertRow(1, 1)

foreach ($col in @(2, 4, 6, 8, 10, 12, 14)) {
    $sheet1.InsertColumn($col, 1)
    $sheet1.Column($col).width = .75
}

Set-ExcelRange -Address $sheet1.Cells["E:E"] -Width 12
Set-ExcelRange -Address $sheet1.Cells["I:I"] -Width 12

$BorderBottom = "Thick"
$BorderColor = "Black"

Set-ExcelRange -Address $sheet1.Cells["A2"]    -BorderBottom $BorderBottom -BorderColor $BorderColor

Set-ExcelRange -Address $sheet1.Cells["C2"]    -BorderBottom $BorderBottom -BorderColor $BorderColor
Set-ExcelRange -Address $sheet1.Cells["E2:G2"] -BorderBottom $BorderBottom -BorderColor $BorderColor
Set-ExcelRange -Address $sheet1.Cells["I2:K2"] -BorderBottom $BorderBottom -BorderColor $BorderColor
Set-ExcelRange -Address $sheet1.Cells["M2:O2"] -BorderBottom $BorderBottom -BorderColor $BorderColor

Set-ExcelRange -Address $sheet1.Cells["A2:C8"] -FontColor Gray

$HorizontalAlignment = "Center"
Set-ExcelRange -Address $sheet1.Cells["F1"] -HorizontalAlignment $HorizontalAlignment -Bold -Value Revenue
Set-ExcelRange -Address $sheet1.Cells["J1"] -HorizontalAlignment $HorizontalAlignment -Bold -Value Margin
Set-ExcelRange -Address $sheet1.Cells["N1"] -HorizontalAlignment $HorizontalAlignment -Bold -Value Passenger

Set-ExcelRange -Address $sheet1.Cells["E2"] -Value '($)'
Set-ExcelRange -Address $sheet1.Cells["G2"] -Value '%'
Set-ExcelRange -Address $sheet1.Cells["I2"] -Value '($)'
Set-ExcelRange -Address $sheet1.Cells["K2"] -Value '%'

Set-ExcelRange -Address $sheet1.Cells["C10"] -HorizontalAlignment Right -Bold -Value "Grand Total Calculation"
Set-ExcelRange -Address $sheet1.Cells["E10"] -Formula "=Sum(E3:E8)" -Bold
Set-ExcelRange -Address $sheet1.Cells["I10"] -Formula "=Sum(I3:I8)" -Bold
Set-ExcelRange -Address $sheet1.Cells["M10"] -Formula "=Sum(M3:M8)" -Bold
Set-ExcelRange -Address $sheet1.Cells["O10"] -Formula "=Sum(O3:O8)" -Bold

Close-ExcelPackage $excel -Show
