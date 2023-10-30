try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$path =  "$env:TEMP\testBackgroundColor.xlsx"

$p = Get-Process | Select-Object Company, Handles | Export-Excel $path -ClearSheet  -PassThru

$ws        = $p.Workbook.WorkSheets[1]
$totalRows = $ws.Dimension.Rows

#Set the range from B2 to the last active row. s
Set-ExcelRange -Range $ws.Cells["B2:B$($totalRows)"] -BackgroundColor LightBlue

Export-Excel -ExcelPackage $p -show -AutoSize