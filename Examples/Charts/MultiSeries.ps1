try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

Remove-Item temp.xlsx -ErrorAction Ignore

$data = invoke-sum (Get-Process) company handles,pm,VirtualMemorySize

$c = New-ExcelChart -Title Stats `
    -ChartType LineMarkersStacked `
    -Header "Stuff" `
    -XRange "Processes[Company]" `
    -YRange "Processes[PM]","Processes[VirtualMemorySize]"

$data |
    Export-Excel temp.xlsx -AutoSize -TableName Processes -Show -ExcelChartDefinition $c
