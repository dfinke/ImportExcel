Get-Process |
    Select-Object Company,Handles,PM, NPM| 
    Export-Excel $xlfile -Show  -AutoSize -CellStyleSB {
        param(
            $workSheet,
            $totalRows,
            $lastColumn
        )
                
        Set-CellStyle $workSheet 1 $LastColumn Solid Cyan

        foreach($row in (2..$totalRows | Where-Object {$_ % 2 -eq 0})) {
            Set-CellStyle $workSheet $row $LastColumn Solid Gray
        }

        foreach($row in (2..$totalRows | Where-Object {$_ % 2 -eq 1})) {
            Set-CellStyle $workSheet $row $LastColumn Solid LightGray
        }
    }