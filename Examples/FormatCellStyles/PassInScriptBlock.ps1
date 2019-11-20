try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$xlfile = "$env:temp\testFmt.xlsx"
Remove-Item $xlfile -ErrorAction Ignore

$RandomStyle = {
    param(
        $workSheet,
        $totalRows,
        $lastColumn
    )

    2..$totalRows | ForEach-Object{
        Set-CellStyle $workSheet $_ $LastColumn Solid (Get-Random @("LightGreen", "Gray", "Red"))
    }
}

Get-Process |
    Select-Object Company,Handles,PM, NPM|
    Export-Excel $xlfile -Show  -AutoSize -CellStyleSB $RandomStyle
