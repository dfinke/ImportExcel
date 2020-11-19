try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

$data = Invoke-Sum -data (Get-Process) -dimension Company -measure Handles, PM, VirtualMemorySize

$c = New-ExcelChartDefinition -Title "ProcessStats" `
    -ChartType LineMarkersStacked `
    -XRange "Processes[Name]" `
    -YRange "Processes[PM]","Processes[VirtualMemorySize]" `
    -SeriesHeader "PM","VM"

$data |
    Export-Excel -Path $xlSourcefile -AutoSize -TableName Processes -ExcelChartDefinition $c  -Show
