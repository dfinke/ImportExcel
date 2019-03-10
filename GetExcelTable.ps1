function Get-ExcelTable {
    <#
    .SYNOPSIS

    .DESCRIPTION
    #>
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        $FullName,
        $WorkSheetName,
        $TableName
    )

    process {
        if (!$WorkSheetName) {$WorkSheetName = "*"}
        if (!$TableName) {$TableName = "*"}

        $excel = Open-ExcelPackage $FullName

        $h = [ordered]@{}

        foreach ($ws in ($excel.Workbook.Worksheets | Where-Object {$_.Name -like $WorkSheetName})) {

            foreach ($table in ($ws.tables | Where-Object {$_.Name -like $TableName})) {
                if (!$h.Contains($ws.Name)) {
                    $h.($ws.Name) = [ordered]@{}
                }

                $s = ""

                $StartRow = $table.Address.Start.Row
                $StartColumn = $table.Address.Start.Column
                $EndRow = $table.Address.End.Row
                $EndColumn = $table.Address.End.Column

                foreach ($row in $StartRow..$EndRow) {
                    $newRow = @()
                    foreach ($column in $StartColumn..$EndColumn) {
                        $newRow += $ws.Cells[$row, $column].value
                    }
                    $s += ($newRow -join ",") + "`r`n"
                }

                $h.($ws.Name).($table.name) = ConvertFrom-Csv $s
            }
        }

        Close-ExcelPackage $excel -NoSave

        $h
    }
}

Set-Alias gxlt Get-ExcelTable