try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

ConvertFrom-ExcelToSQLInsert People .\testSQLGen.xlsx

ConvertFrom-ExcelData .\testSQLGen.xlsx {
    param($propertyNames, $record)

    $reportRecord = @()
    foreach ($pn in $propertyNames) {
        $reportRecord += "{0}: {1}" -f $pn, $record.$pn
    }
    $reportRecord +=""
    $reportRecord -join "`r`n"
}