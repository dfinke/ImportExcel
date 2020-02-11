$xlfile = "$env:TEMP\Points.xlsx"

Remove-Item $xlfile -ErrorAction SilentlyContinue

$PivotTableDefinition = New-PivotTableDefinition -Activate -PivotTableName Points `
    -PivotRows Driver, Date -PivotData @{Points = "SUM"} -GroupDateRow Date -GroupDatePart Years, Months

Import-Csv "$PSScriptRoot\First10Races.csv" |
    Select-Object  Race, @{n = "Date"; e = {[datetime]::ParseExact($_.date, "dd/MM/yyyy", (Get-Culture))}}, FinishPosition, Driver, GridPosition, Team, Points |
    Export-Excel $xlfile -Show -AutoSize -PivotTableDefinition $PivotTableDefinition