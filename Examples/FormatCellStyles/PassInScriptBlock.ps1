$RandomStyle = {
    param(
        $workSheet,
        $totalRows,
        $lastColumn
    )    

    2..$totalRows | ForEach-Object{
        Set-CellStyle $workSheet $_ $LastColumn Solid (Write-Output LightGreen Gray Red|Get-Random)
    }
}

Get-Process |
    Select-Object Company,Handles,PM, NPM| 
    Export-Excel $xlfile -Show  -AutoSize -CellStyleSB $RandomStyle
