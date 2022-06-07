function Enable-ExcelAutoFilter {
    <#
        .SYNOPSIS
        Enable the Excel AutoFilter

        .EXAMPLE
        Enable-ExcelAutoFilter $targetSheet
    #>    
    param(
        [Parameter(Mandatory)]
        [OfficeOpenXml.ExcelWorksheet]$Worksheet 
    )

    $range = Get-ExcelSheetDimensionAddress $Worksheet    
    $Worksheet.Cells[$range].AutoFilter = $true
}