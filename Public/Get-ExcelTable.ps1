function Get-ExcelTable {
    <#
        .Synopsis
        .Example
    #>

    param(
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        $Path,
        $WorksheetName
    )

    Process {
        try {
            $Error.Clear()
            $pkg = Open-ExcelPackage -Path $Path
        }
        catch {
             "$Path - $($Error.exception.InnerException.message)"
        }
        $ws = $pkg.Workbook.Worksheets

        $result = foreach ($table in $ws.Tables) {
            [PSCustomObject][Ordered]@{
                Path          = $Path
                WorksheetName = $table.WorkSheet
                TableName     = $table.Name
                Address       = $table.Address
                Columns       = $table.Columns
            }
        }

    }
    
    End {
        if ($pkg) {
            Close-ExcelPackage -ExcelPackage $pkg -NoSave

            if ($WorksheetName) {
                $result | Where-Object { $_.WorksheetName -like $WorksheetName }
            }
            else {
                $result
            }
        }
    }
}