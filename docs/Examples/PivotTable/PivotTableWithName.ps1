try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$ExcelParams = @{
    Path              = "$env:TEMP\test1.xlsx"
    IncludePivotTable = $true
    PivotRows         = 'Company'
    PivotTableName    = 'MyTable'
    PivotData         = @{'Handles' = 'sum'}
    Show              = $true
    Activate          = $true
}
Remove-Item $ExcelParams.Path -ErrorAction Ignore
Get-Process | Select-Object Company, Handles | Export-Excel @ExcelParams

<#  Builds a pivot table that looks like this:

    Sum of Handles
    Row Labels	                 Total
    Adobe Systems Incorporated	  3100
    (blank)	                    214374
    Apple Inc.	                   215
    etc
    etc
    Grand Total	                365625
#>