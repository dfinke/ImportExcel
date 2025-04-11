try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$xlSourcefile = "$env:TEMP\NoFormulaExample.xlsx"

#Remove existing file
Remove-Item $xlSourcefile -ErrorAction Ignore

#Simulate a matching comparison to get a SideIndicator of '=='
$comparison = Compare-Object -ReferenceObject @(1) -DifferenceObject @(1) -IncludeEqual

#Add '-NoFormulaConversion SideIndicator' to allow '==' to be used in the SideIndicator column without converting to a formula
$comparison | Export-Excel $xlSourcefile -NoFormulaConversion SideIndicator -Show
