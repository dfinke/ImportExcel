function New-ExcelChartDefinition {
    [Alias("New-ExcelChart")] #This was the former name. The new name reflects that we are defining a chart, not making one in the workbook.
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Does not change system State')]
    param(
        $Title = "Chart Title",
        $Header,
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = "ColumnStacked",
        [OfficeOpenXml.Drawing.Chart.eTrendLine[]]$ChartTrendLine,
        $XRange,
        $YRange,
        $Width = 500,
        $Height = 350,
        $Row = 0,
        $RowOffSetPixels = 10,
        $Column = 6,
        $ColumnOffSetPixels = 5,
        [OfficeOpenXml.Drawing.Chart.eLegendPosition]$LegendPosition,
        $LegendSize,
        [Switch]$LegendBold,
        [Switch]$NoLegend,
        [Switch]$ShowCategory,
        [Switch]$ShowPercent,
        $SeriesHeader,
        [Switch]$TitleBold,
        [Int]$TitleSize ,
        [String]$XAxisTitleText,
        [Switch]$XAxisTitleBold,
        $XAxisTitleSize ,
        [string]$XAxisNumberformat,
        $XMajorUnit,
        $XMinorUnit,
        $XMaxValue,
        $XMinValue,
        [OfficeOpenXml.Drawing.Chart.eAxisPosition]$XAxisPosition        ,
        [String]$YAxisTitleText,
        [Switch]$YAxisTitleBold,
        $YAxisTitleSize,
        [string]$YAxisNumberformat,
        $YMajorUnit,
        $YMinorUnit,
        $YMaxValue,
        $YMinValue,
        [OfficeOpenXml.Drawing.Chart.eAxisPosition]$YAxisPosition
    )
    if ( $Header ) { Write-Warning "The header parameter is ignored." } #Nothing was done with it when creating a chart.
    #might be able to do [PSCustomObject]$PsboundParameters,   the defaults here match those in Add-Excel Chart
    [PSCustomObject]@{
        Title              = $Title
        ChartType          = $ChartType
        ChartTrendLine     = $ChartTrendLine
        XRange             = $XRange
        YRange             = $YRange
        Width              = $Width
        Height             = $Height
        Row                = $Row
        RowOffSetPixels    = $RowOffSetPixels
        Column             = $Column
        ColumnOffSetPixels = $ColumnOffSetPixels
        LegendPosition     = $LegendPosition
        LegendSize         = $LegendSize
        Legendbold         = $LegendBold
        NoLegend           = $NoLegend -as [Boolean]
        ShowCategory       = $ShowCategory -as [Boolean]
        ShowPercent        = $ShowPercent -as [Boolean]
        SeriesHeader       = $SeriesHeader
        TitleBold          = $TitleBold -as [Boolean]
        TitleSize          = $TitleSize
        XAxisTitleText     = $XAxisTitleText
        XAxisTitleBold     = $XAxisTitleBold -as [Boolean]
        XAxisTitleSize     = $XAxisTitleSize
        XAxisNumberformat  = $XAxisNumberformat
        XMajorUnit         = $XMajorUnit
        XMinorUnit         = $XMinorUnit
        XMaxValue          = $XMaxValue
        XMinValue          = $XMinValue
        XAxisPosition      = $XAxisPosition
        YAxisTitleText     = $YAxisTitleText
        YAxisTitleBold     = $YAxisTitleBold -as [Boolean]
        YAxisTitleSize     = $YAxisTitleSize
        YAxisNumberformat  = $YAxisNumberformat
        YMajorUnit         = $YMajorUnit
        YMinorUnit         = $YMinorUnit
        YMaxValue          = $YMaxValue
        YMinValue          = $YMinValue
        YAxisPosition      = $YAxisPosition
    }
}
