try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

$PivotTableDefinition = New-PivotTableDefinition -Activate -PivotTableName Places `
    -PivotRows Driver, FinishPosition -PivotData @{Date = "Count"} -GroupNumericRow FinishPosition -GroupNumericMin 1 -GroupNumericMax 25 -GroupNumericInterval 3

Import-Csv "$PSScriptRoot\First10Races.csv" |
    Select-Object  Race, @{n = "Date"; e = {[datetime]::ParseExact($_.date, "dd/MM/yyyy", (Get-Culture))}}, FinishPosition, Driver, GridPosition, Team, Points |
    Export-Excel $xlSourcefile -Show -AutoSize -PivotTableDefinition $PivotTableDefinition