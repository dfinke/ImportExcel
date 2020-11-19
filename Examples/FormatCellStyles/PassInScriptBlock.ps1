try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

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
    Export-Excel $xlSourcefile -Show  -AutoSize -CellStyleSB $RandomStyle
