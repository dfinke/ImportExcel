function New-ExcelChart {
    param(
        $Title="Chart Title",
        $Header,
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType="ColumnStacked",
        $XRange,
        $YRange,
        $Width=500,
        $Height=350,
        $Row=0,
        $RowOffSetPixels=10,
        $Column=6,
        $ColumnOffSetPixels=5,
        [Switch]$NoLegend,
        [Switch]$ShowCategory,
        [Switch]$ShowPercent,
        $SeriesHeader
    )

    [PSCustomObject]@{
        Title=$Title
        Header=$Header
        ChartType=$ChartType
        XRange=$XRange
        YRange=$YRange
        Width=$Width
        Height=$Height
        Row=$Row
        RowOffSetPixels=$RowOffSetPixels
        Column=$Column
        ColumnOffSetPixels=$ColumnOffSetPixels
        NoLegend     = if($NoLegend)     {$true} else {$false}
        ShowCategory = if($ShowCategory) {$true} else {$false}
        ShowPercent  = if($ShowPercent)  {$true} else {$false}
        SeriesHeader=$SeriesHeader
   }
}