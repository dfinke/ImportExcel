try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

ColumnChart -Title "Central Limit Theorem" -NoLegend ($(
        for ($i = 1; $i -le 500; $i++) {
            $s = 0
            for ($j = 1; $j -le 100; $j++) {
                $s += Get-Random -Minimum 0 -Maximum 2
            }
            $s
        }
    ) | Sort-Object | Group-Object | Select-Object Count, Name)