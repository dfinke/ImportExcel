function Get-XYRange {
    [CmdletBinding()]
    param($TargetData)

    $record = $TargetData | Select-Object -First 1
    $p=$record.psobject.Properties.name

    $infer = for ($idx = 0; $idx -lt $p.Count; $idx++) {

        $name = $p[$idx]
        $value = $record.$name

        $result=Invoke-AllTests $value -OnlyPassing -FirstOne

        [PSCustomObject]@{
            Name         = $name
            Value        = $value
            DataType     = $result.DataType
            ExcelColumn  = (Get-ExcelColumnName ($idx+1)).ColumnName
        }
    }

    [PSCustomObject]@{
        XRange = $infer | Where-Object -FilterScript {$_.datatype -match 'string'} | Select-Object -First 1 -Property excelcolumn, name
        YRange = $infer | Where-Object -FilterScript {$_.datatype -match 'int|double'} | Select-Object -First 1 -Property excelcolumn, name
    }
}