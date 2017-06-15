Clear-Host

Import-Module .\ImportExcel.psd1 -Force

#ConvertFrom-ExcelToSQLInsert People .\testSQLGen.xlsx

ConvertFrom-ExcelData .\testSQLGen.xlsx {
    param($propertyNames, $record)

    $reportRecord = @()
    foreach ($pn in $propertyNames) {
        $reportRecord += "{0}: {1}" -f $pn, $record.$pn
    }
    $reportRecord +=""
    $reportRecord -join "`r`n"
}

return 

ConvertFrom-ExcelData .\testSQLGen.xlsx {
    param($propertyNames, $record)

    $ColumnNames = "'" + ($PropertyNames -join "', '") + "'"
    $values = foreach ($propertyName in $PropertyNames) { $record.$propertyName }
    $targetValues = "'" + ($values -join "', '") + "'"

    "INSERT INTO People ({0}) Values({1});" -f $ColumnNames, $targetValues
}