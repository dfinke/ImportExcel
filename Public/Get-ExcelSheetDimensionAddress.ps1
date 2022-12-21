function Get-ExcelSheetDimension {
    <#
        .SYNOPSIS
        Get the Excel address of the dimension of a sheet

        .PARAMETER
        Object of [OfficeOpenXml.ExcelWorksheet]

        .EXAMPLE
        $excelPackage=Open-ExcelPackage -Path "H:\Test\test.xlsx"
        Get-ExcelSheetDimensionAddress -Worksheet $excelPackage.Workbook.Worksheets[1]
        Close-ExcelPackage -ExcelPackage $excelPackage

        .OUTPUT
        RANGE        : I9:AJ78
        END_Column   : 36
        END_Row      : 78
        START_Column : 9
        START_Row    : 9
    #>
    
    param(
        [Parameter(Mandatory)]
        [OfficeOpenXml.ExcelWorksheet]$Worksheet
    )
    $dimensionProperties=$Worksheet.Dimension
    [PSCustomObject]@{
        RANGE = $dimensionProperties.Address
        END_Column = $dimensionProperties.End.Column
        END_Row = $dimensionProperties.End.Row
        START_Column = $dimensionProperties.Start.Column
        START_Row = $dimensionProperties.Start.Column
    }
}
