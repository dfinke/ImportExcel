$xlfile = "$env:TEMP\Places.xlsx"

Remove-Item $xlfile -ErrorAction SilentlyContinue

$PivotTableDefinition = New-PivotTableDefinition -Activate -PivotTableName Places `
    -PivotRows Driver, FinishPosition -PivotData @{Date = "Count"} -GroupNumericRow FinishPosition -GroupNumericMin 1 -GroupNumericMax 25 -GroupNumericInterval 3

Import-Csv "$PSScriptRoot\First10Races.csv" |
    Select-Object  Race, @{n = "Date"; e = {[datetime]::ParseExact($_.date, "dd/MM/yyyy", (Get-Culture))}}, FinishPosition, Driver, GridPosition, Team, Points |
    Export-Excel $xlfile -Show -AutoSize -PivotTableDefinition $PivotTableDefinition