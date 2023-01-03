
function Add-ExcelChart {
    <#
    .SYNOPSIS
    Add Excel Chart into excel sheet
    
    .DESCRIPTION
    Function is creating chart by data passed in parameter
    
    .PARAMETER Worksheet
    Handler to worksheet object
    
    .PARAMETER PivotTable
    Handler to pivot table (created before)
    
    .PARAMETER Title
    Title of excel chart
    
    .PARAMETER ChartType
    Type of chart. For more info just open excel and check rule how it works :)

    .PARAMETER ChartTrendLine
    Trend line used in Excel to indicate trend in chart. Like trend in getting raised salary by month
    
    .PARAMETER XRange
    Range of values passed into X-Axis. On Chart visible horizontally.
    
    .PARAMETER YRange
    Range of values passed into Y-Axis. On Chart visible perpendicuraly.
    
    .PARAMETER Width
    Width of Chart
    
    .PARAMETER Height
    Height of Chart
    
    .PARAMETER Row
    Number of row where left upper corner of chart is started
    
    .PARAMETER RowOffSetPixels
    Parameter description
    
    .PARAMETER Column
    Number of column where left upper corner of chart is started
    
    .PARAMETER ColumnOffSetPixels
    Parameter description
    
    .PARAMETER LegendPosition
    Position of legend 
    
    .PARAMETER LegendSize
    Size of text used in legend
    
    .PARAMETER LegendBold
    Bold text in a legend
    
    .PARAMETER NoLegend
    Do not show legend on chart
    
    .PARAMETER ShowCategory
    Parameter description
    
    .PARAMETER ShowPercent
    Parameter description
    
    .PARAMETER SeriesHeader
    Parameter description
    
    .PARAMETER TitleBold
    When true then chart title is bold
    
    .PARAMETER TitleSize
    Size of text title of chart
    
    .PARAMETER XAxisTitleText
    Title of XAxis
    
    .PARAMETER XAxisTitleBold
    When true then XAxis title is bold
    
    .PARAMETER XAxisTitleSize
    Size of text title of XAxis
    
    .PARAMETER XAxisNumberformat
    Number format used on XAxis. For more info just open excel and check rule how it works :)
    
    .PARAMETER XMajorUnit
    Main unit used on XAxis
    
    .PARAMETER XMinorUnit
    Second unit used on XAxis
    
    .PARAMETER XMaxValue
    Max value used on XAxis
    
    .PARAMETER XMinValue
    Min value used on XAxis
    
    .PARAMETER XAxisPosition
    Position of XAxis. For more info just open excel and check rule how it works :)
    
    .PARAMETER YAxisTitleText
    Title of YAxis
    
    .PARAMETER YAxisTitleBold
    When true then YAxis title is bold
    
    .PARAMETER YAxisTitleSize
    Size of text title of YAxis
    
    .PARAMETER YAxisNumberformat
    Number format used on YAxis. For more info just open excel and check rule how it works :)
    
    .PARAMETER YMajorUnit
    Main unit used on YAxis
    
    .PARAMETER YMinorUnit
    Second unit used on YAxis
    
    .PARAMETER YMaxValue
    Maximum value of YAxis.
    
    .PARAMETER YMinValue
    Minimum value of YAxis.
    
    .PARAMETER YAxisPosition
    Position of YAxis. For more info just open excel and check rule how it works :)
    
    .PARAMETER PassThru
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    General notes
    #>
    [CmdletBinding(DefaultParameterSetName = 'Worksheet')]
    [OutputType([OfficeOpenXml.Drawing.Chart.ExcelChart])]
    param(
        [Parameter(ParameterSetName = 'Worksheet', Mandatory = $true)]
        [OfficeOpenXml.ExcelWorksheet]$Worksheet,
        [Parameter(ParameterSetName = 'PivotTable', Mandatory = $true)]
        [OfficeOpenXml.Table.PivotTable.ExcelPivotTable]$PivotTable ,
        [String]$Title,
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = "ColumnStacked",
        [OfficeOpenXml.Drawing.Chart.eTrendLine[]]$ChartTrendLine,
        $XRange,
        $YRange,
        [int]$Width = 500,
        [int]$Height = 350,
        [int]$Row = 0,
        [int]$RowOffSetPixels = 10,
        [int]$Column = 6,
        [int]$ColumnOffSetPixels = 5,
        [OfficeOpenXml.Drawing.Chart.eLegendPosition]$LegendPosition,
        $LegendSize,
        [Switch]$LegendBold,
        [Switch]$NoLegend,
        [Switch]$ShowCategory,
        [Switch]$ShowPercent,
        [String[]]$SeriesHeader,
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
            $Worksheet = $PivotTable.Worksheet
            $chart = $Worksheet.Drawings.AddChart(("Chart" + $PivotTable.Name ), $ChartType, $PivotTable)
        }
        else {
            $ChartName = 'Chart' + (Split-Path -Leaf ([System.IO.path]::GetTempFileName())) -replace 'tmp|\.', ''
            $chart = $Worksheet.Drawings.AddChart($ChartName, $ChartType)
            $chartDefCount = @($YRange).Count
            if ($chartDefCount -eq 1) {
                $Series = $chart.Series.Add($YRange, $XRange)
                if ($ChartTrendLine) {
                    if ($ChartType -notmatch "stacked|3D$|pie|Doughnut|Cone|Cylinder|Pyramid") {
                        foreach ($trendLine in $ChartTrendLine) {
                            $null = $Series.TrendLines.Add($trendLine)
                        }
                    }
                    else {
                        Write-Warning "Chart trend line is not supported for chart type: $ChartType"
                    }
                }
                if ($SeriesHeader) { $Series.Header = $SeriesHeader }
                else { $Series.Header = 'Series 1' }
            }
            else {
                for ($idx = 0; $idx -lt $chartDefCount; $idx += 1) {
                    if ($Yrange.count -eq $xrange.count) {
                        $Series = $chart.Series.Add($YRange[$idx], $XRange[$idx])
                    }
                    else {
                        $Series = $chart.Series.Add($YRange[$idx], $XRange)
                    }
                    if ($SeriesHeader.Count -gt 0) {
                        if ($SeriesHeader[$idx] -match '^=') { $Series.HeaderAddress = $SeriesHeader[$idx] -replace '^=', '' }
                        else { $Series.Header = $SeriesHeader[$idx] }
                    }
                    else { $Series.Header = "Series $($idx)" }
                }
            }
        }
        if ($Title) {
            $chart.Title.Text = $Title
            if ($TitleBold) { $chart.Title.Font.Bold = $true }
            if ($TitleSize) { $chart.Title.Font.Size = $TitleSize }
        }
        if ($NoLegend) { $chart.Legend.Remove() }
        else {
            if ($PSBoundParameters.ContainsKey('LegendPosition')) { $chart.Legend.Position = $LegendPosition }
            if ($PSBoundParameters.ContainsKey('LegendBold')) { $chart.Legend.Font.Bold = [boolean]$LegendBold }
            if ($LegendSize) { $chart.Legend.Font.Size = $LegendSize }
        }

        if ($XAxisTitleText) {
            $chart.XAxis.Title.Text = $XAxisTitleText
            if ($PSBoundParameters.ContainsKey('XAxisTitleBold')) {
                $chart.XAxis.Title.Font.Bold = [boolean]$XAxisTitleBold
            }
            if ($XAxisTitleSize) { $chart.XAxis.Title.Font.Size = $XAxisTitleSize }
        }
        if ($XAxisPosition) { Write-Warning "X-axis position is not being set propertly at the moment, parameter ignored" }
        #$chart.ChartXml.chartSpace.chart.plotArea.catAx.axPos.val = $XAxisPosition.ToString().substring(0,1)}
        if ($XMajorUnit) { $chart.XAxis.MajorUnit = $XMajorUnit }
        if ($XMinorUnit) { $chart.XAxis.MinorUnit = $XMinorUnit }
        if ($null -ne $XMinValue) { $chart.XAxis.MinValue = $XMinValue }
        if ($null -ne $XMaxValue) { $chart.XAxis.MaxValue = $XMaxValue }
        if ($XAxisNumberformat) { $chart.XAxis.Format = (Expand-NumberFormat $XAxisNumberformat) }

        if ($YAxisTitleText) {
            $chart.YAxis.Title.Text = $YAxisTitleText
            if ($PSBoundParameters.ContainsKey('YAxisTitleBold')) {
                $chart.YAxis.Title.Font.Bold = [boolean]$YAxisTitleBold
            }
            if ($YAxisTitleSize) { $chart.YAxis.Title.Font.Size = $YAxisTitleSize }
        }
        if ($YAxisPosition) { Write-Warning "Y-axis position is not being set propertly at the moment, parameter ignored" }
        #$chart.ChartXml.chartSpace.chart.plotArea.valAx.axPos.val= $YAxisPosition.ToString().substring(0,1)}
        if ($YMajorUnit) { $chart.YAxis.MajorUnit = $YMajorUnit }
        if ($YMinorUnit) { $chart.YAxis.MinorUnit = $YMinorUnit }
        if ($null -ne $YMinValue) { $chart.YAxis.MinValue = $YMinValue }
        if ($null -ne $YMaxValue) { $chart.YAxis.MaxValue = $YMaxValue }
        if ($YAxisNumberformat) { $chart.YAxis.Format = (Expand-NumberFormat $YAxisNumberformat) }
        if ($null -ne $chart.Datalabel) {
            $chart.Datalabel.ShowCategory = [boolean]$ShowCategory
            $chart.Datalabel.ShowPercent = [boolean]$ShowPercent
        }

        $chart.SetPosition($Row, $RowOffsetPixels, $Column, $ColumnOffsetPixels)
        $chart.SetSize($Width, $Height)

        if ($PassThru) { return $chart }
    }
    catch { Write-Warning -Message "Failed adding Chart to worksheet '$($Worksheet).name': $_" }
}
