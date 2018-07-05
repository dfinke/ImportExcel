function Get-XYRange {
    param($targetData)

    $record = $targetData| Select-Object -First 1
    $p=$record.psobject.Properties.name

    $infer = for ($idx = 0; $idx -lt $p.Count; $idx++) {

        $name = $p[$idx]
        $value = $record.$name

        $result=Invoke-AllTests $value -OnlyPassing -FirstOne

        [PSCustomObject]@{
            Name         = $name
            Value        = $value
            DataType     = $result.DataType
            ExcelColumn = [OfficeOpenXml.ExcelAddress]::TranslateFromR1C1("R[1]C[$($idx+1)]", 0 , 0) -replace "\d+", ""  #(Get-ExcelColumnName ($idx + 1)).ColumnName
        }
    }

    [PSCustomObject]@{
        XRange = $infer | ? {$_.datatype -match 'string'} | Select-Object -First 1 excelcolumn, name
        YRange = $infer | ? {$_.datatype -match 'int|double'} |Select-Object -First 1 excelcolumn, name
    }
}