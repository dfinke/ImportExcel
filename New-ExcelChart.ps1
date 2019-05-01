function New-ExcelChartDefinition {
    <#
      .SYNOPSIS
        Creates a Definition of a chart which can be added using Export-Excel, or Add-PivotTable
      .DESCRIPTION
        All the parameters which are passed to Add-ExcelChart can be added to an object and
        passed to Export-Excel with the -ExcelChartDefinition parameter,
        or to Add-PivotTable with the -PivotChartDefinition parameter.
        This command sets up those definitions.
      .PARAMETER Title
        The title for the chart.
      .PARAMETER TitleBold
        Sets the title in bold face.
      .PARAMETER TitleSize
        Sets the point size for the title.
      .PARAMETER ChartType
        One of the built-in chart types, such as Pie, ClusteredColumn, Line etc. Defaults to "ColumnStacked".
      .PARAMETER XRange
        The range of cells containing values for the X-Axis - usually labels.
      .PARAMETER YRange
        The range(s) of cells holding values for the Y-Axis - usually "data".
      .PARAMETER Width
        Width of the chart in pixels. Defaults to 500.
      .PARAMETER Height
        Height of the chart in pixels. Defaults to 350.
      .PARAMETER Row
        Row position of the top left corner of the chart. 0 places at the top of the sheet, 1 below row 1 and so on.
      .PARAMETER RowOffSetPixels
        Offset to position the chart by a fraction of a row.
      .PARAMETER Column
        Column position of the top left corner of the chart. 0 places at the edge of the sheet 1 to the right of column A and so on.
      .PARAMETER ColumnOffSetPixels
        Offset to position the chart by a fraction of a column.
      .PARAMETER NoLegend
        If specified, turns off display of the key. If you only have one data series it may be preferable to use the title to say what the chart is.
      .PARAMETER SeriesHeader
        Specifies explicit name(s) for the data series, which will appear in the legend/key
      .PARAMETER LegendPosition
        Location of the key, either "Left", "Right", "Top", "Bottom" or "TopRight".
      .PARAMETER LegendSize
        Font size for the key.
      .PARAMETER LegendBold
        Sets the key in bold type.
      .PARAMETER ShowCategory
        Attaches a category label in charts which support this.
      .PARAMETER ShowPercent
        Attaches a percentage label in charts which support this.
      .PARAMETER XAxisTitleText
        Specifies a title for the X-axis.
      .PARAMETER XAxisTitleBold
        Sets the X-axis title in bold face.
      .PARAMETER XAxisTitleSize
        Sets the font size for the axis title.
      .PARAMETER XAxisNumberformat
        A number formatting string, like "#,##0.00", for numbers along the X-axis.
      .PARAMETER XMajorUnit
        Spacing for the major gridlines / tick marks along the X-axis.
      .PARAMETER XMinorUnit
        Spacing for the minor gridlines / tick marks along the X-axis.
      .PARAMETER XMaxValue
        Maximum value for the scale along the X-axis.
      .PARAMETER XMinValue
        Minimum value for the scale along the X-axis.
      .PARAMETER xAxisPosition
        Position for the X-axis ("Top" or" Bottom").
      .PARAMETER YAxisTitleText
        Specifies a title for the Y-axis.
      .PARAMETER YAxisTitleBold
        Sets the Y-axis title in bold face.
      .PARAMETER YAxisTitleSize
        Sets the font size for the Y-axis title.
      .PARAMETER YAxisNumberformat
        A number formatting string, like "#,##0.00", for numbers on the Y-axis
      .PARAMETER YMajorUnit
        Spacing for the major gridlines / tick marks on the Y-axis.
      .PARAMETER YMinorUnit
        Spacing for the minor gridlines / tick marks on the Y-axis.
      .PARAMETER YMaxValue
        Maximum value on the Y-axis.
      .PARAMETER YMinValue
        Minimum value on the Y-axis.
      .PARAMETER YAxisPosition
        Position for the Y-axis ("Left" or "Right").
      .PARAMETER Header
        No longer used. This may be removed in future versions.
      .Example
        >
        PS> $cDef = New-ExcelChartDefinition  -ChartType line -XRange "X" -YRange "Sinx"  -Title "Graph of Sine X" -TitleBold -TitleSize 14 `
                       -Column 2 -ColumnOffSetPixels 35 -Width 800 -XAxisTitleText "Degrees" -XAxisTitleBold -XAxisTitleSize 12 -XMajorUnit 30 -XMinorUnit 10 -XMinValue 0 -XMaxValue 361  -XAxisNumberformat "000" `
                       -YMinValue -1.25 -YMaxValue 1.25 -YMajorUnit 0.25 -YAxisNumberformat "0.00" -YAxisTitleText "Sine" -YAxisTitleBold -YAxisTitleSize 12 `
                       -SeriesHeader "Sin(x)" -LegendSize 8 -legendBold  -LegendPosition Bottom

        0..360 | ForEach-Object {[pscustomobject][ordered]@{x = $_; Sinx = "=Sin(Radians(x)) "}} | Export-Excel -AutoNameRange -now -WorkSheetname SinX -ExcelChartDefinition $cDef -Show

        This reworks an example from Add-Excel-Chart but here the chart is defined
        and the defintion stored in $cDef and then Export-Excel uses $cDef .
    #>
    [Alias("New-ExcelChart")] #This was the former name. The new name reflects that we are defining a chart, not making one in the workbook.
    [cmdletbinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',Justification='Does not change system State')]
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
        LegendPosition      = $LegendPosition
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
      .SYNOPSIS
        Creates a chart in an existing Excel worksheet.
      .DESCRIPTION
        Creates a chart. It is possible to configure the type of chart, the range of X values (labels) and Y values.
        the title, the legend, the ranges for both axes, the format and position of the axes.
        Normally the command does not return anything, but if -passthru is specified the chart is returned so that it can be customized.
      .PARAMETER Worksheet
        An existing Sheet where the chart will be created.
      .PARAMETER Title
        The title for the chart.
      .PARAMETER TitleBold
        Sets the title in bold face.
      .PARAMETER TitleSize
        Sets the point size for the title.
      .PARAMETER ChartType
        One of the built-in chart types, such as Pie, ClusteredColumn, Line etc. Defaults to "ColumnStacked".
      .PARAMETER XRange
        The range of cells containing values for the X-Axis - usually labels.
       .PARAMETER YRange
        The range(s) of cells holding values for the Y-Axis - usually "data".
      .PARAMETER PivotTable
        Instead of specify X and Y ranges, get data from a PivotTable by passing a PivotTable Object.
      .PARAMETER Width
        Width of the chart in Pixels. Defaults to 500.
      .PARAMETER Height
        Height of the chart in Pixels. Defaults to 350.
      .PARAMETER Row
        Row position of the top left corner of the chart. 0 places at the top of the sheet, 1 below row 1 and so on.
      .PARAMETER RowOffSetPixels
        Offset to position the chart by a fraction  of a row.
      .PARAMETER Column
        Column position of the top left corner of the chart. 0 places at the edge of the sheet 1 to the right of column A and so on.
      .PARAMETER ColumnOffSetPixels
        Offset to position the chart by a fraction of a column.
      .PARAMETER NoLegend
        If specified, turns of display of the key. If you only have one data series it may be preferable to use the title to say what the chart is.
      .PARAMETER SeriesHeader
        Specify explicit name(s) for the data series, which will appear in the legend/key. The contents of a cell can be specified in the from =Sheet9!Z10 .
       .PARAMETER LegendPosition
        Location of the key, either left, right, top, bottom or TopRight.
      .PARAMETER LegendSize
        Font size for the key.
      .PARAMETER LegendBold
        Sets the key in bold type.
      .PARAMETER ShowCategory
        Attaches a category label in charts which support this.
      .PARAMETER ShowPercent
        Attaches a percentage label in charts which support this.
      .PARAMETER XAxisTitleText
        Specifies a title for the X-axis.
      .PARAMETER XAxisTitleBold
        Sets the X-axis title in bold face.
      .PARAMETER XAxisTitleSize
        Sets the font size for the axis title.
      .PARAMETER XAxisNumberformat
        A number formatting string, like "#,##0.00", for numbers along the X-axis.
      .PARAMETER XMajorUnit
        Spacing for the major gridlines / tick marks along the X-axis.
      .PARAMETER XMinorUnit
        Spacing for the minor gridlines / tick marks along the X-axis.
      .PARAMETER XMaxValue
        Maximum value for the scale along the X-axis.
      .PARAMETER XMinValue
        Minimum value for the scale along the X-axis.
      .PARAMETER xAxisPosition
        Position for the X-axis (Top or Bottom).
      .PARAMETER YAxisTitleText
        Specifies a title for the Y-axis.
      .PARAMETER YAxisTitleBold
        Sets the Y-axis title in bold face.
      .PARAMETER YAxisTitleSize
        Sets the font size for the Y-axis title
      .PARAMETER YAxisNumberformat
        A number formatting string, like "#,##0.00", for numbers on the Y-axis.
      .PARAMETER YMajorUnit
        Spacing for the major gridlines / tick marks on the Y-axis.
      .PARAMETER YMinorUnit
        Spacing for the minor gridlines / tick marks on the Y-axis.
      .PARAMETER YMaxValue
        Maximum value on the Y-axis.
      .PARAMETER YMinValue
        Minimum value on the Y-axis.
      .PARAMETER YAxisPosition
        Position for the Y-axis (Left or Right).
      .PARAMETER PassThru
        Add-Excel chart doesn't normally return anything, but if -PassThru is specified it returns the newly created chart to allow it to be fine tuned.
      .EXAMPLE
        >
        PS> $Excel = ConvertFrom-Csv    @"
        Product, City, Sales
        Apple, London , 300
        Orange, London , 400
        Banana, London , 300
        Orange, Paris,   600
        Banana, Paris,   300
        Apple, New York, 1200
        "@  | Export-Excel  -Path test.xlsx -PassThru
        Add-ExcelChart -Worksheet $Excel.Workbook.Worksheets[1] -ChartType "Doughnut" -XRange "A2:B7"  -YRange "C2:C7" -width 500
        Close-ExcelPackage -Show $Excel

        The first command expands a multi-line string into 6 rows of data which is exported to new Excel file; leaving an ExcelPackage object in $excel
        The second command adds a chart - the cell ranges are explicitly specified. Note that the XRange (labels) is TWO columns wide and the chart will
        combine the name of the product and the name of the City to create the label.
        The width of the chart is set explictly, the default legend is used and there is no Chart title.
      .EXAMPLE
        >
        PS> $Excel = Invoke-Sum (Get-Process) Company Handles, PM, VirtualMemorySize | Export-Excel $path  -AutoSize -ExcelChartDefinition $c -AutoNameRange -PassThru
        Add-ExcelChart -Worksheet $Excel.Workbook.Worksheets[1] -Title "VM use" -ChartType PieExploded3D   -XRange "Name" -YRange "VirtualMemorySize" -NoLegend -ShowCategory
        Close-ExcelPackage $Excel -Show

        The first line exports information and creates named ranges for each column.
        The Second line uses the ranges to specify a chart - the labels come from the range "Name" and the data from the range "VirtualMemorySize"
        The chart is specified as a 3D exploded PIE chart, with a title of "VM Use" and instead of a legend the the pie slices are identified with a label.
      .EXAMPLE
        >
        PS> $Excel = Invoke-Sum (Get-Process) Company Handles, PM, VirtualMemorySize | Export-Excel test.xlsx  -TableName Processes -PassThru
        Add-ExcelChart -Worksheet $Excel.Workbook.Worksheets[1] -Title Stats -ChartType LineMarkersStacked   -XRange "Processes[Name]" -YRange "Processes[PM]", "Processes[VirtualMemorySize]" -SeriesHeader 'PM', 'VMSize'
        Close-ExcelPackage $Excel -Show

        The first line exports information to a table in new file; and captures the excel Package object in $Excel
        The second line creates a chart on the first page of the work sheet, using the notation "TableName[ColumnnName]" to refer to the data,
        the labels come Name column in the table, and the data series from its PM and VirtualMemorySize columns. The display names for these in the header are set to PM and VMSize
      .EXAMPLE
        >
        PS> $excel = 0..360 | ForEach-Object {[pscustomobject][ordered]@{x = $_; Sinx = "=Sin(Radians(x)) "}} | Export-Excel -AutoNameRange -Path Text.xlsx -WorkSheetname SinX -PassThru
        Add-ExcelChart -Worksheet $excel.Workbook.Worksheets["Sinx"] -ChartType line -XRange "X" -YRange "Sinx"  -Title "Graph of Sine X" -TitleBold -TitleSize 14 `
                       -Column 2 -ColumnOffSetPixels 35 -Width 800 -XAxisTitleText "Degrees" -XAxisTitleBold -XAxisTitleSize 12 -XMajorUnit 30 -XMinorUnit 10 -XMinValue 0 -XMaxValue 361  -XAxisNumberformat "000" `
                       -YMinValue -1.25 -YMaxValue 1.25 -YMajorUnit 0.25 -YAxisNumberformat "0.00" -YAxisTitleText "Sine" -YAxisTitleBold -YAxisTitleSize 12 `
                       -SeriesHeader "Sin(x)" -LegendSize 8 -legendBold  -LegendPosition Bottom
        Close-ExcelPackage $Excel -Show

        The first line puts numbers from 0 to 360 into a sheet, as the first column, and
        a formula to calculate the Sine of that number of number of degrees in the second column.
        It creates named-ranges for the two columns - "X" and "SinX" respectively
        The Add-ExcelChart command adds a chart to that worksheet, specifying a line chart
        with the X values coming from named-range "X" and the Y values coming from named-range "SinX".
        The chart has a title, and is positioned to the right of column 2 and sized 800 pixels wide
        The X-axis is labelled "Degrees", in bold 12 point type and runs from 0 to 361 with labels every 30,
        and minor tick marks every 10. Degrees are shown padded to 3 digits.
        The Y-axis is labelled "Sine" and to allow some room above and below its scale runs from -1.25 to 1.25,
        and is marked off in units of 0.25 shown to two decimal places.
        The key will for the chart will be at the bottom in 8 point bold type and the line will be named "Sin(x)".
    #>
    [cmdletbinding(DefaultParameterSetName='Worksheet')]
    [OutputType([OfficeOpenXml.Drawing.Chart.ExcelChart])]
    param(
        [Parameter(ParameterSetName='Workshet',Mandatory=$true)]
        [OfficeOpenXml.ExcelWorksheet]$Worksheet,
        [Parameter(ParameterSetName='PivotTable',Mandatory=$true)]
        [OfficeOpenXml.Table.PivotTable.ExcelPivotTable]$PivotTable ,
        [String]$Title,
        #$Header,   Not used but referenced previously
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = "ColumnStacked",
        $XRange,
        $YRange,
        [int]$Width              = 500,
        [int]$Height             = 350,
        [int]$Row                =   0,
        [int]$RowOffSetPixels    =  10,
        [int]$Column             =   6,
        [int]$ColumnOffSetPixels =   5,
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
                    if ($Yrange.count -eq $xrange.count) {
                      $Series = $chart.Series.Add($YRange[$idx], $XRange[$idx])
                    }
                    else {
                      $Series = $chart.Series.Add($YRange[$idx], $XRange)
                    }
                    if ($SeriesHeader.Count -gt 0) {
                      if ($SeriesHeader[$idx] -match '^=') {$Series.HeaderAddress = $SeriesHeader[$idx] -replace '^=',''}
                      else                                 {$Series.Header        = $SeriesHeader[$idx] }
                    }
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
            if ($PSBoundParameters.ContainsKey('LegendPosition')) {$chart.Legend.Position  = $LegendPosition}
            if ($PSBoundParameters.ContainsKey('LegendBold'))     {$chart.Legend.Font.Bold = [boolean]$LegendBold}
            if ($LegendSize)                                      {$chart.Legend.Font.Size = $LegendSize}
        }

        if ($XAxisTitleText)      {
            $chart.XAxis.Title.Text = $XAxisTitleText
            if ($PSBoundParameters.ContainsKey('XAxisTitleBold'))  {
                                   $chart.XAxis.Title.Font.Bold = [boolean]$XAxisTitleBold
            }
            if ($XAxisTitleSize)  {$chart.XAxis.Title.Font.Size = $XAxisTitleSize}
        }
        if ($XAxisPosition)       {Write-Warning "X-axis position is not being set propertly at the moment, parameter ignored" }
                                   #$chart.ChartXml.chartSpace.chart.plotArea.catAx.axPos.val = $XAxisPosition.ToString().substring(0,1)}
        if ($XMajorUnit)          {$chart.XAxis.MajorUnit       = $XMajorUnit}
        if ($XMinorUnit)          {$chart.XAxis.MinorUnit       = $XMinorUnit}
        if ($null -ne $XMinValue) {$chart.XAxis.MinValue        = $XMinValue}
        if ($null -ne $XMaxValue) {$chart.XAxis.MaxValue        = $XMaxValue}
        if ($XAxisNumberformat)   {$chart.XAxis.Format          = (Expand-NumberFormat $XAxisNumberformat)}

       if ($YAxisTitleText)     {
            $chart.YAxis.Title.Text = $YAxisTitleText
            if ($PSBoundParameters.ContainsKey('YAxisTitleBold')) {
                                  $chart.YAxis.Title.Font.Bold = [boolean]$YAxisTitleBold
            }
            if ($YAxisTitleSize) {$chart.YAxis.Title.Font.Size = $YAxisTitleSize}
        }
        if ($YAxisPosition)      {Write-Warning "Y-axis position is not being set propertly at the moment, parameter ignored" }
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
