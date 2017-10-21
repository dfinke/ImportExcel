
function Import-ExcelTable
{
    [CmdletBinding()]
    Param (
        [Parameter(Position=1, Mandatory)]
        [String]$Path,
        [Parameter(Position=2, Mandatory)]
        [Alias('Sheet')]
        [String]$WorksheetName,
        [Parameter(Position=3, Mandatory)]
        [Alias('Table')]
        [String]$TableName,
        [ValidateRange(1, 9999)]
        [Int]$TopRow
    )

$Path = (Resolve-Path $Path).ProviderPath
Write-Verbose "Import Excel workbook '$Path' with worksheet '$Worksheetname'"

$Stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Path, 'Open', 'Read', 'ReadWrite'
$Excel = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Stream

        
if (-not ($Worksheet = $Excel.Workbook.Worksheets[$WorkSheetName])) {
    throw "Worksheet '$WorksheetName' not found, the workbook only contains the worksheets '$($Excel.Workbook.Worksheets)'."
}
if (-not ($excelTable = $Worksheet.Tables[$TableName])) {
    throw "Table '$TableName' not found in the worksheet. Worksheet only contains the tables '$($Worsheet.Tables)'."
}
    $rows = @()
    $excelTable = $Worksheet.Tables[$TableName]
    $StartRow = $excelTable.Address.Start.Row + 1
    $StartColumn = $excelTable.Address.Start.Column
    $RowCount = $excelTable.Address.Rows - 2
           
    if($TopRow -and $RowCount -gt $TopRow)
    {
         $RowCount = ($TopRow - 1)
    }

    $ColumnCount = $excelTable.Address.Columns

    $EndRow = $StartRow + $RowCount
            
    $EndColumn = $StartColumn + $ColumnCount
        
    foreach($Row in $StartRow..($EndRow))
    {
        $newRow = [Ordered]@{}
        $CellsWithValues = $worksheet.Cells[$Row, $StartColumn, $Row, $EndColumn] | Where-Object Value
        if($CellsWithValues)
        {
            foreach($Column in $excelTable.Columns)
            {
                $propertyName = $Column.Name
                $position = $Column.Position
                $newRow."$propertyName" = $worksheet.Cells[($Row),($position+$StartColumn)].Value
            }
            $rows += [PSCustomObject]$newRow
        }
    }
    $rows 
}    

   
