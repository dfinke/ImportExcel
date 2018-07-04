function New-ExcelChartDefinition {
    [cmdletbinding()]
    [Alias("New-ExcelChart")] #This was the former name. The new name reflects that we are defining a chart, not making one in the workbook. 
    param(
        $Title = "Chart Title",
        $Header,
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = "ColumnStacked",
        $XRange,
        $YRange,
        $Width = 500,
        $Height = 350,
        $Row = 0,
        $RowOffSetPixels = 10,
        $Column = 6,
        $ColumnOffSetPixels = 5,
        [Switch]$NoLegend,
        [Switch]$ShowCategory,
        [Switch]$ShowPercent,
        $SeriesHeader
    )
    if ( $Header ) {Write-Warning "The header parameter is ignored."} #Nothing was done with it when creating a chart. 
    [PSCustomObject]@{
        Title              = $Title
        ChartType          = $ChartType
        XRange             = $XRange
        YRange             = $YRange
        Width              = $Width
        Height             = $Height
        Row                = $Row
        RowOffSetPixels    = $RowOffSetPixels
        Column             = $Column
        ColumnOffSetPixels = $ColumnOffSetPixels
        NoLegend           = $NoLegend    -as [Boolean]
        ShowCategory       = $ShowCategory-as [Boolean]
        ShowPercent        = $ShowPercent -as [Boolean]
        SeriesHeader       = $SeriesHeader
    }
}