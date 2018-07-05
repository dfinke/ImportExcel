try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

Remove-Item .\test1.xlsx -ErrorAction Ignore

$ExcelParams = @{
    Path              = ".\test1.xlsx"
    IncludePivotTable = $true
    PivotRows         = 'Company'
    PivotTableName    = 'MyTable'
    PivotData         = @{'Handles' = 'sum'}
    Show              = $true
}

Get-Process | Select-Object Company, Handles |
    Export-Excel @ExcelParams