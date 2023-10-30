function Enable-ExcelAutofit {
    <#
        .SYNOPSIS
        Make all text fit the cells
        
        .EXAMPLE
        Enable-ExcelAutofit $excel.Sheet1
    #>    
    param(
        [Parameter(Mandatory)]
        [OfficeOpenXml.ExcelWorksheet]$Worksheet 
    )

    $range = Get-ExcelSheetDimensionAddress $Worksheet
    $Worksheet.Cells[$range].AutoFitColumns()
}