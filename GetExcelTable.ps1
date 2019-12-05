function Get-ExcelTableName {
    param(
        $Path,
        $WorksheetName
    )

    $Path = (Resolve-Path $Path).ProviderPath
    $Stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Path, 'Open', 'Read', 'ReadWrite'

    $Excel = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Stream

    if ($WorksheetName) {
        $Worksheet = $Excel.Workbook.Worksheets[$WorkSheetName]
    } else {
        $Worksheet = $Excel.Workbook.Worksheets | Select-Object -First 1
    }

    foreach($TableName in $Worksheet.Tables.Name) {
        [PSCustomObject][Ordered]@{
            WorksheetName=$Worksheet.Name
            TableName=$TableName
        }
    }

    $Stream.Close()
    $Stream.Dispose()
    $Excel.Dispose()
    $Excel = $null
}

function Get-ExcelTable {
    param(
        $Path,
        $TableName,
        $WorksheetName
    )

    $Path = (Resolve-Path $Path).ProviderPath
    $Stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Path, 'Open', 'Read', 'ReadWrite'

    $Excel = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Stream

    if ($WorksheetName) {
        $Worksheet = $Excel.Workbook.Worksheets[$WorkSheetName]
    } else {
        $Worksheet = $Excel.Workbook.Worksheets | Select-Object -First 1
    }

    if($TableName) {
        $Table = $Worksheet.Tables[$TableName]
    } else {
        $Table = $Worksheet.Tables | Select-Object -First 1
    }

    $rowCount = $Table.Address.Rows
    $colCount = $Table.Address.Columns

    $digits = "0123456789".ToCharArray()

    $start, $end=$Table.Address.Address.Split(':')

    $pos=$start.IndexOfAny($digits)
    [int]$startCol=ConvertFrom-ExcelColumnName $start.Substring(0,$pos)
    [int]$startRow=$start.Substring($pos)

    $propertyNames = for($col=$startCol; $col -lt ($startCol+$colCount); $col+= 1) {
        $Worksheet.Cells[$startRow, $col].value
    }

    $startRow++
    for($row=$startRow; $row -lt ($startRow+$rowCount); $row += 1) {
        $nr=[ordered]@{}
        $c=0
        for($col=$startCol; $col -lt ($startCol+$colCount); $col+= 1) {
            $nr.($propertyNames[$c]) = $Worksheet.Cells[$row, $col].value
            $c++
        }
        [pscustomobject]$nr
    }

    $Stream.Close()
    $Stream.Dispose()
    $Excel.Dispose()
    $Excel = $null
}

function ConvertFrom-ExcelColumnName {
    param($columnName)

    $sum=0
    $columnName.ToCharArray() |
        ForEach-Object {
            $sum*=26
            $sum+=[char]$_.tostring().toupper()-[char]'A'+1
        }
    $sum
}

Import-Module .\ImportExcel.psd1 -Force

#Get-ExcelTableName .\testTable.xlsx | Get-ExcelTable .\testTable.xlsx
Get-ExcelTable .\testTable.xlsx Table3