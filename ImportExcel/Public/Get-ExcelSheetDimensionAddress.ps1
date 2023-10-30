function Get-ExcelSheetDimensionAddress {
    <#
        .SYNOPSIS
        Get the Excel address of the dimension of a sheet

        .EXAMPLE
        Get-ExcelSheetDimensionAddress $excel.Sheet1
    #>
    param(
        [Parameter(Mandatory)]
        [OfficeOpenXml.ExcelWorksheet]$Worksheet 
    )

    $Worksheet.Dimension.Address
}
