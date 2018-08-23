function New-ExcelChartDefinition {
    [Alias("New-ExcelChart")] #This was the former name. The new name reflects that we are defining a chart, not making one in the workbook.
    [cmdletbinding()]
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
        [OfficeOpenXml.Drawing.Chart.eLegendPosition]$LegendPostion,
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
    if ( $Header ) {Write-Warning "The header parameter is ignored."} #Nothing was done with it when creating a chart.
    #might be able to do [PSCustomObject]$PsboundParameters,   the defaults here match those in Add-Excel Chart
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
        LegendPostion      = $LegendPostion
        LegendSize         = $LegendSize
        Legendbold         = $LegendBold
        NoLegend           = $NoLegend     -as [Boolean]
        ShowCategory       = $ShowCategory -as [Boolean]
        ShowPercent        = $ShowPercent  -as [Boolean]
        SeriesHeader       = $SeriesHeader
        TitleBold          = $TitleBold    -as [Boolean]
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
        YAxisTitleBold     = $YAxisTitleBold  -as [Boolean]
        YAxisTitleSize     = $YAxisTitleSize
        YAxisNumberformat  = $YAxisNumberformat
        YMajorUnit         = $YMajorUnit
        YMinorUnit         = $YMinorUnit
        YMaxValue          = $YMaxValue
        YMinValue          = $YMinValue
        YAxisPosition      = $YAxisPosition
    }
}

function Add-ExcelChart {
    <#
      .Synopsis
        Creates a chart in an Existing excel worksheet
    #>
    [cmdletbinding(DefaultParameterSetName='Worksheet')]
    param(
        #An object representing the worksheet where the chart should be added.
        [Parameter(ParameterSetName='Workshet',Mandatory=$true)]
        [OfficeOpenXml.ExcelWorksheet]$Worksheet,
        [Parameter(ParameterSetName='PivotTable',Mandatory=$true)]
        [OfficeOpenXml.Table.PivotTable.ExcelPivotTable]$PivotTable ,
        [String]$Title = "Chart Title",
        #$Header,   Not used but referenced previously
        #The Type of chart (Area, Line, Pie etc)
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = "ColumnStacked",
        $XRange,
        $YRange,
        [int]$Width              = 500,
        [int]$Height             = 350,
        [int]$Row                =   0,
        [int]$RowOffSetPixels    =  10,
        [int]$Column             =   6,
        [int]$ColumnOffSetPixels =   5,
        [OfficeOpenXml.Drawing.Chart.eLegendPosition]$LegendPostion,
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
        [OfficeOpenXml.Drawing.Chart.eAxisPosition]$YAxisPosition,
        [Switch]$PassThru
        )
    try {
        if ($PivotTable) {
            $Worksheet = $PivotTable.WorkSheet
            $chart     = $Worksheet.Drawings.AddChart(("Chart" + $PivotTable.Name ),$ChartType,$PivotTable)
        }
        else {
            $ChartName = 'Chart' + (Split-Path -Leaf ([System.IO.path]::GetTempFileName())) -replace 'tmp|\.', ''
            $chart = $Worksheet.Drawings.AddChart($ChartName, $ChartType)
            $chartDefCount = @($YRange).Count
            if ($chartDefCount -eq 1) {
                $Series = $chart.Series.Add($YRange, $XRange)
                if ($SeriesHeader) { $Series.Header = $SeriesHeader}
                else { $Series.Header = 'Series 1'}
            }
            else {
                for ($idx = 0; $idx -lt $chartDefCount; $idx += 1) {
                    $Series = $chart.Series.Add($YRange[$idx], $XRange)
                    if ($SeriesHeader.Count -gt 0) { $Series.Header = $SeriesHeader[$idx] }
                    else { $Series.Header = "Series $($idx)"}
                }
            }
        }
        if ($Title) {
            $chart.Title.Text = $Title
            if ($TitleBold) {$chart.Title.Font.Bold = $true}
            if ($TitleSize) {$chart.Title.Font.Size = $TitleSize}
        }
        if ($NoLegend) { $chart.Legend.Remove() }
        else {
            if ($LegendPostion) {$Chart.Legend.Position    = $LegendPostion}
            if ($LegendSize)    {$chart.Legend.Font.Size   = $LegendSize}
            if ($legendBold)    {$chart.Legend.Font.Bold   = $true}
        }

        if ($XAxisTitleText)      {
            $chart.XAxis.Title.Text = $XAxisTitleText
            if ($XAxisTitleBold)  {$chart.XAxis.Title.Font.Bold = $true}
            if ($XAxisTitleSize)  {$chart.XAxis.Title.Font.Size = $XAxisTitleSize}
        }
        if ($XAxisPosition)       {Write-Warning "X Axis position is not being set propertly at the moment, parameter ignored" }
                                   #$chart.ChartXml.chartSpace.chart.plotArea.catAx.axPos.val = $XAxisPosition.ToString().substring(0,1)}
        if ($XMajorUnit)          {$chart.XAxis.MajorUnit       = $XMajorUnit}
        if ($XMinorUnit)          {$chart.XAxis.MinorUnit       = $XMinorUnit}
        if ($null -ne $XMinValue) {$chart.XAxis.MinValue        = $XMinValue}
        if ($null -ne $XMaxValue) {$chart.XAxis.MaxValue        = $XMaxValue}
        if ($XAxisNumberformat)   {$chart.XAxis.Format          = (Expand-NumberFormat $XAxisNumberformat)}

       if ($YAxisTitleText)     {
            $chart.YAxis.Title.Text = $YAxisTitleText
            if ($YAxisTitleBold) {$chart.YAxis.Title.Font.Bold = $true}
            if ($YAxisTitleSize) {$chart.YAxis.Title.Font.Size = $YAxisTitleSize}
        }
        if ($YAxisPosition)      {Write-Warning "Y Axis position is not being set propertly at the moment, parameter ignored" }
                                  #$chart.ChartXml.chartSpace.chart.plotArea.valAx.axPos.val= $YAxisPosition.ToString().substring(0,1)}
        if ($YMajorUnit)         {$chart.YAxis.MajorUnit       = $YMajorUnit}
        if ($YMinorUnit)         {$chart.YAxis.MinorUnit       = $YMinorUnit}
        if ($null -ne $YMinValue){$chart.YAxis.MinValue        = $YMinValue}
        if ($null -ne $YMaxValue){$chart.YAxis.MaxValue        = $YMaxValue}
        if ($YAxisNumberformat)  {$chart.YAxis.Format          = (Expand-NumberFormat $YAxisNumberformat)}
        if ($null -ne $chart.Datalabel) {
                                  $chart.Datalabel.ShowCategory = [boolean]$ShowCategory
                                  $chart.Datalabel.ShowPercent  = [boolean]$ShowPercent
        }

        $chart.SetPosition($Row, $RowOffsetPixels, $Column, $ColumnOffsetPixels)
        $chart.SetSize($Width, $Height)


        if ($PassThru) {return $chart}
    }
    catch {Write-Warning -Message "Failed adding Chart to worksheet '$($WorkSheet).name': $_"}
}
