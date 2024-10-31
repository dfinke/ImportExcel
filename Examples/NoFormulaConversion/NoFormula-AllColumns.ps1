try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$xlSourcefile = "$env:TEMP\NoFormulaExample.xlsx"

#Remove existing file
Remove-Item $xlSourcefile -ErrorAction Ignore

#These formulas should not calculate and remain as text
[PSCustOmobject][Ordered]@{   
    Formula1 = "=COUNT(1,1,1)"
    Formula2 = "=SUM(2,3)"
    Formula3 = "=1=1"
} | Export-Excel $xlSourcefile -NoFormulaConversion * -Calculate -Show
