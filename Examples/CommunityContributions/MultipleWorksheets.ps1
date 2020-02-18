<#
    To see this written up with example screenshots, head over to the IT Splat blog
    URL: http://bit.ly/2SxieeM
#>

## Create an Excel file with multiple worksheets
# Get a list of processes on the system
$processes = Get-Process | Sort-Object -Property ProcessName | Group-Object -Property ProcessName | Where-Object {$_.Count -gt 2}

# Export the processes to Excel, each process on its own sheet
$processes | ForEach-Object { $_.Group | Export-Excel -Path MultiSheetExample.xlsx -WorksheetName $_.Name -AutoSize -AutoFilter }

# Show the completed file
Invoke-Item .\MultiSheetExample.xlsx

## Add an additional sheet to the new workbook
# Use Open-ExcelPackage to open the workbook
$excelPackage = Open-ExcelPackage -Path .\MultiSheetExample.xlsx

# Create a new worksheet and give it a name, set MoveToStart to make it the first sheet
$ws = Add-Worksheet -ExcelPackage $excelPackage -WorksheetName 'All Services' -MoveToStart

# Get all the running services on the system
Get-Service | Export-Excel -ExcelPackage $excelPackage -WorksheetName $ws -AutoSize -AutoFilter

# Close the package and show the final result
Close-ExcelPackage -ExcelPackage $excelPackage -Show
