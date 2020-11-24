
Function Import-Bycolumns    {
    Param(
        [Parameter(Mandatory=$true)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        [Int]$StartRow = 1,
        [String]$WorksheetName,
        [Int]$EndRow ,
        [Int]$StartColumn = 1,
        [Int]$EndColumn
    )
    Function Get-RowNames {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "Name would be incorrect, and command is not exported")]
        param(
            [Parameter(Mandatory)]
            [Int[]]$Rows,
            [Parameter(Mandatory)]
            [Int]$StartColumn
        )
        foreach ($R in $Rows) {
            #allow "False" or "0" to be  headings
            $Worksheet.Cells[$R, $StartColumn] | Where-Object {-not [string]::IsNullOrEmpty($_.Value) } | Select-Object @{N = 'Row'; E = { $R } }, Value
        }
    }

    if     (-not  $WorksheetName) { $Worksheet = $ExcelPackage.Workbook.Worksheets[1] }
    elseif (-not ($Worksheet = $ExcelPackage.Workbook.Worksheets[$WorkSheetName])) {
        throw "Worksheet '$WorksheetName' not found, the workbook only contains the worksheets '$($ExcelPackage.Workbook.Worksheets)'. If you only wish to select the first worksheet, please remove the '-WorksheetName' parameter." ; return
    }

    if (-not $EndRow   ) { $EndRow     = $Worksheet.Dimension.End.Row }
    if (-not $EndColumn) { $EndColumn = $Worksheet.Dimension.End.Column }

    $Rows    = $Startrow .. $EndRow  ;
    $Columns = (1 + $StartColumn)..$EndColumn

    if ((-not $rows) -or (-not ($PropertyNames = Get-RowNames -Rows $Rows -StartColumn $StartColumn))) {
        throw "No headers found in left coulmn '$Startcolumn'. "; return
    }
    if (-not $Columns) {
        Write-Warning "Worksheet '$WorksheetName' in workbook contains no data in the rows after left column '$StartColumn'"
    }
    else {
        foreach ($c in $Columns) {
            $NewColumn = [Ordered]@{ }
            foreach ($p in $PropertyNames) {
                $NewColumn[$p.Value] = $Worksheet.Cells[$p.row,$c].text
            }
            [PSCustomObject]$NewColumn
        }
    }
}
