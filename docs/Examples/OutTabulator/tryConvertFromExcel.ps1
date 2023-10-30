[CmdletBinding()]
param($outFile = "$PSScriptRoot\targetout.html")

$columnOptions = @()

$columnOptions += New-ColumnOption -ColumnName Progress -formatter progress
$columnOptions += New-ColumnOption -ColumnName Activity -formatter lineFormatter

ConvertFrom-Excel -ExcelFile $PSScriptRoot\test.xlsx -outFile $PSScriptRoot\targetout.html -columnOptions $columnOptions