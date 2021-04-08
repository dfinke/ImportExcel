function Add-ExcelTableData {
    <#
        .Synopsis
        .Example
    #>
    param(
        [Parameter(ValueFromPipelineByPropertyName)]
        $Path,
        [Parameter(ValueFromPipelineByPropertyName)]
        $WorksheetName,
        [Parameter(ValueFromPipelineByPropertyName)]
        $TableName,
        $Data
    )

    Process {
        $pkg = Open-ExcelPackage -Path $Path
        $ws = $pkg.Workbook.Worksheets[$WorksheetName]
        $targetTable = $ws.Tables[$TableName]
        if ($null -eq $targetTable) {
            Write-Warning -Message "The Table '$TableName' was not found."
        }
        else {
            $startColumn = $targetTable.Address.Start.Column
            $row = $targetTable.Address.End.Row + 1
            $names = $data[0].psobject.Properties.Name
    
            foreach ($record in $data) {    
                foreach ($name in $names) {
                    $targetColumn = $targetTable.Columns[$name]
                    $column = $startColumn + $targetColumn.Position
                    if (!$targetColumn) {
                        Write-Warning -Message "The Column name '$name' was not found."
                        continue
                    }
                    else {
                        $ws.Cells[$row, $column].Value = $record.$name
                    }
                }
                $row += 1
            }

            # $newRange = Add-NumRowsToRange $ws.Dimension.Address ($Data.count - 1)
            # $newRange = Add-NumRowsToRange $targetTable.Address.Address ($Data.count - 1)
            $newRange = Add-NumRowsToRange $targetTable.Address.Address $Data.count 

            # $targetTable.TableXml.table.ref = $newRange

            $targetTable.TableXml.table.ref = $ws.Cells[$newRange].Address
            #$targetTable.TableXml.table.ref = $ws.Cells[$ws.Dimension].Address
        }

        Close-ExcelPackage -ExcelPackage $pkg 
    }
}