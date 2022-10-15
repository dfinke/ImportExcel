function Get-ExcelFileSummary {
    <#
        .Synopsis
        Gets summary information on an Excel file like number of rows, columns, and more
    #>
    param(
        [Parameter(ValueFromPipelineByPropertyName, Mandatory)]
        [Alias('FullName')]    
        $Path
    )

    Process {    
        $excel = Open-ExcelPackage -Path $Path

        foreach ($workSheet in $excel.Workbook.Worksheets) {        
            [PSCustomObject][Ordered]@{
                ExcelFile     = Split-Path -Leaf $Path
                WorksheetName = $workSheet.Name
                Visible       = $workSheet.Hidden -eq 'Visible'
                Rows          = $workSheet.Dimension.Rows
                Columns       = $workSheet.Dimension.Columns
                Address       = $workSheet.Dimension.Address
                Path          = Split-Path  $Path
            }
        }

        Close-ExcelPackage -ExcelPackage $excel -NoSave
    }
}