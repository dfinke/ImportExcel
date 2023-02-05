---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Add-ExcelChart

## SYNOPSIS

Creates a chart in an existing Excel worksheet.

## SYNTAX

### Worksheet \(Default\)

```text
Add-ExcelChart [-Title <String>] [-ChartType <eChartType>] [-ChartTrendLine <eTrendLine[]>] [-XRange <Object>] [-YRange <Object>] [-Width <Int32>] [-Height <Int32>] [-Row <Int32>] [-RowOffSetPixels <Int32>]  [-Column <Int32>] [-ColumnOffSetPixels <Int32>] [-LegendPosition <eLegendPosition>] [-LegendSize <Object>] [-LegendBold] [-NoLegend] [-ShowCategory] [-ShowPercent] [-SeriesHeader <String[]>] [-TitleBold] [-TitleSize <Int32>] [-XAxisTitleText <String>] [-XAxisTitleBold] [-XAxisTitleSize <Object>] [-XAxisNumberformat <String>] [-XMajorUnit <Object>] [-XMinorUnit <Object>] [-XMaxValue <Object>] [-XMinValue <Object>] [-XAxisPosition <eAxisPosition>] [-YAxisTitleText <String>] [-YAxisTitleBold] [-YAxisTitleSize <Object>] [-YAxisNumberformat <String>] [-YMajorUnit <Object>] [-YMinorUnit <Object>] [-YMaxValue <Object>] [-YMinValue <Object>] [-YAxisPosition <eAxisPosition>] [-PassThru] [<CommonParameters>]
```

### PivotTable

```text
Add-ExcelChart -PivotTable <ExcelPivotTable> [-Title <String>] [-ChartType <eChartType>] [-ChartTrendLine <eTrendLine[]>] [-XRange <Object>] [-YRange <Object>] [-Width <Int32>] [-Height <Int32>] [-Row <Int32>] [-RowOffSetPixels <Int32>] [-Column <Int32>] [-ColumnOffSetPixels <Int32>] [-LegendPosition <eLegendPosition>] [-LegendSize <Object>] [-LegendBold] [-NoLegend] [-ShowCategory] [-ShowPercent] [-SeriesHeader <String[]>] [-TitleBold] [-TitleSize <Int32>] [-XAxisTitleText <String>]  [-XAxisTitleBold] [-XAxisTitleSize <Object>] [-XAxisNumberformat <String>] [-XMajorUnit <Object>] [-XMinorUnit <Object>] [-XMaxValue <Object>] [-XMinValue <Object>] [-XAxisPosition <eAxisPosition>] [-YAxisTitleText <String>] [-YAxisTitleBold] [-YAxisTitleSize <Object>] [-YAxisNumberformat <String>] [-YMajorUnit <Object>] [-YMinorUnit <Object>] [-YMaxValue <Object>] [-YMinValue <Object>]  [-YAxisPosition <eAxisPosition>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION

Creates a chart.

It is possible to configure the type of chart, the range of X values \(labels\) and Y values, the title, the legend, the ranges for both axes, the format and position of the axes.

Normally the command does not return anything, but if -passthru is specified the chart is returned so that it can be customized.

## EXAMPLES

### EXAMPLE 1

```text
PS\> $Excel = ConvertFrom-Csv    @"
    Product, City,    Sales
    Apple,   London,    300
    Orange,  London,    400
    Banana,  London,    300
    Orange,  Paris,     600
    Banana,  Paris,     300
    Apple,   New York, 1200
"@  | Export-Excel  -Path test.xlsx -PassThru
    Add-ExcelChart -Worksheet $Excel.Workbook.Worksheets[1] -ChartType "Doughnut" -XRange "A2:B7"  -YRange "C2:C7" -width 500
    Close-ExcelPackage -Show $Excel
```

The first command expands a multi-line string into 6 rows of data which is exported to new Excel file; leaving an ExcelPackage object in $excel The second command adds a chart - the cell ranges are explicitly specified.

Note that the XRange \(labels\) is TWO columns wide and the chart will combine the name of the product and the name of the City to create the label.

The width of the chart is set explictly, the default legend is used and there is no Chart title.

### EXAMPLE 2

```text
PS\> $Excel = Invoke-Sum (Get-Process) Company Handles, PM, VirtualMemorySize | Export-Excel $path  -AutoSize -ExcelChartDefinition $c -AutoNameRange -PassThru
     Add-ExcelChart -Worksheet $Excel.Workbook.Worksheets[1] -Title "VM use" -ChartType PieExploded3D   -XRange "Name" -YRange "VirtualMemorySize" -NoLegend -ShowCategory
     Close-ExcelPackage $Excel -Show
```

The first line exports information and creates named ranges for each column.

The Second line uses the ranges to specify a chart - the labels come from the range "Name" and the data from the range "VirtualMemorySize"

The chart is specified as a 3D exploded PIE chart, with a title of "VM Use" and instead of a legend the the pie slices are identified with a label.

### EXAMPLE 3

```text
PS\> $Excel = Invoke-Sum (Get-Process) Company Handles, PM, VirtualMemorySize | Export-Excel test.xlsx  -TableName Processes -PassThru
      Add-ExcelChart -Worksheet $Excel.Workbook.Worksheets[1] -Title Stats -ChartType LineMarkersStacked   -XRange "Processes[Name]" -YRange "Processes[PM]", "Processes[VirtualMemorySize]" -SeriesHeader 'PM', 'VMSize'
      Close-ExcelPackage $Excel -Show
```

The first line exports information to a table in new file; and captures the excel Package object in $Excel

The second line creates a chart on the first page of the work sheet, using the notation "TableName\[ColumnName\]" to refer to the data, the labels come Name column in the table, and the data series from its PM and VirtualMemorySize columns. The display names for these in the header are set to 'PM' and 'VMSize'.

### EXAMPLE 4

```text
PS\> $excel = 0..360 | ForEach-Object {[pscustomobject][ordered]@{x = $_; Sinx = "=Sin(Radians(x)) "}} | Export-Excel -AutoNameRange -Path Text.xlsx -WorkSheetname SinX -PassThru
Add-ExcelChart -Worksheet $excel.Workbook.Worksheets["Sinx"] -ChartType line -XRange "X" -YRange "Sinx"  -Title "Graph of Sine X" -TitleBold -TitleSize 14 \`
               -Column 2 -ColumnOffSetPixels 35 -Width 800 -XAxisTitleText "Degrees" -XAxisTitleBold -XAxisTitleSize 12 -XMajorUnit 30 -XMinorUnit 10 -XMinValue 0 -XMaxValue 361  -XAxisNumberformat "000" \`
               -YMinValue -1.25 -YMaxValue 1.25 -YMajorUnit 0.25 -YAxisNumberformat "0.00" -YAxisTitleText "Sine" -YAxisTitleBold -YAxisTitleSize 12 \`
               -SeriesHeader "Sin(x)" -LegendSize 8 -legendBold  -LegendPosition Bottom
Close-ExcelPackage $Excel -Show
```

The first line puts numbers from 0 to 360 into a sheet, as the first column, and a formula to calculate the Sine of that number of number of degrees in the second column. It creates named-ranges for the two columns - "X" and "SinX" respectively

The Add-ExcelChart command adds a chart to that worksheet, specifying a line chart with the X values coming from named-range "X" and the Y values coming from named-range "SinX". The chart has a title, and is positioned to the right of column 2 and sized 800 pixels wide

The X-axis is labelled "Degrees", in bold 12 point type and runs from 0 to 361 with labels every 30, and minor tick marks every 10. Degrees are shown padded to 3 digits.

The Y-axis is labelled "Sine" and to allow some room above and below its scale runs from -1.25 to 1.25, and is marked off in units of 0.25 shown to two decimal places.

The key will for the chart will be at the bottom in 8 point bold type and the line will be named "Sin\(x\)".

## PARAMETERS

### -Worksheet

An existing Sheet where the chart will be created.

```yaml
Type: ExcelWorksheet
Parameter Sets: Worksheet
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PivotTable

Instead of specify X and Y ranges, get data from a PivotTable by passing a PivotTable Object.

```yaml
Type: ExcelPivotTable
Parameter Sets: PivotTable
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title

The title for the chart.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartType

One of the built-in chart types, such as Pie, ClusteredColumn, Line etc. Defaults to "ColumnStacked".

```yaml
Type: eChartType
Parameter Sets: (All)
Aliases:
Accepted values: Area, Line, Pie, Bubble, ColumnClustered, ColumnStacked, ColumnStacked100, ColumnClustered3D, ColumnStacked3D, ColumnStacked1003D, BarClustered, BarStacked, BarStacked100, BarClustered3D, BarStacked3D, BarStacked1003D, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, PieExploded, PieExploded3D, BarOfPie, XYScatterSmooth, XYScatterSmoothNoMarkers, XYScatterLines, XYScatterLinesNoMarkers, AreaStacked, AreaStacked100, AreaStacked3D, AreaStacked1003D, DoughnutExploded, RadarMarkers, RadarFilled, Surface, SurfaceWireframe, SurfaceTopView, SurfaceTopViewWireframe, Bubble3DEffect, StockHLC, StockOHLC, StockVHLC, StockVOHLC, CylinderColClustered, CylinderColStacked, CylinderColStacked100, CylinderBarClustered, CylinderBarStacked, CylinderBarStacked100, CylinderCol, ConeColClustered, ConeColStacked, ConeColStacked100, ConeBarClustered, ConeBarStacked, ConeBarStacked100, ConeCol, PyramidColClustered, PyramidColStacked, PyramidColStacked100, PyramidBarClustered, PyramidBarStacked, PyramidBarStacked100, PyramidCol, XYScatter, Radar, Doughnut, Pie3D, Line3D, Column3D, Area3D

Required: False
Position: Named
Default value: ColumnStacked
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartTrendLine

```yaml
Type: eTrendLine[]
Parameter Sets: (All)
Aliases:
Accepted values: Exponential, Linear, Logarithmic, MovingAvgerage, Polynomial, Power

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -XRange

The range of cells containing values for the X-Axis - usually labels.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -YRange

The range\(s\) of cells holding values for the Y-Axis - usually "data".

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width

Width of the chart in Pixels; defaults to 500.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 500
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height

Height of the chart in Pixels; defaults to 350.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 350
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row

Row position of the top left corner of the chart. \) places at the top of the sheet, 1 below row 1 and so on.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -RowOffSetPixels

Offset to position the chart by a fraction of a row.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 10
Accept pipeline input: False
Accept wildcard characters: False
```

### -Column

Column position of the top left corner of the chart; 0 places at the edge of the sheet 1 to the right of column A and so on.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 6
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnOffSetPixels

Offset to position the chart by a fraction of a column.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 5
Accept pipeline input: False
Accept wildcard characters: False
```

### -LegendPosition

Location of the key, either left, right, top, bottom or TopRight.

```yaml
Type: eLegendPosition
Parameter Sets: (All)
Aliases:
Accepted values: Top, Left, Right, Bottom, TopRight

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LegendSize

Font size for the key.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LegendBold

Sets the key in bold type.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoLegend

If specified, turns of display of the key. If you only have one data series it may be preferable to use the title to say what the chart is.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowCategory

Attaches a category label, in charts which support this.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowPercent

Attaches a percentage label, in charts which support this.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -SeriesHeader

Specify explicit name\(s\) for the data series, which will appear in the legend/key. The contents of a cell can be specified in the from =Sheet9!Z10 .

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TitleBold

Sets the title in bold face.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -TitleSize

Sets the point size for the title.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -XAxisTitleText

Specifies a title for the X-axis.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -XAxisTitleBold

Sets the X-axis title in bold face.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -XAxisTitleSize

Sets the font size for the axis title.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -XAxisNumberformat

A number formatting string, like "\#,\#\#0.00", for numbers along the X-axis.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -XMajorUnit

Spacing for the major gridlines / tick marks along the X-axis.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -XMinorUnit

Spacing for the minor gridlines / tick marks along the X-axis.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -XMaxValue

Maximum value for the scale along the X-axis.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -XMinValue

Minimum value for the scale along the X-axis.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -XAxisPosition

Position for the X-axis \(Top or Bottom\).

```yaml
Type: eAxisPosition
Parameter Sets: (All)
Aliases:
Accepted values: Left, Bottom, Right, Top

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -YAxisTitleText

Specifies a title for the Y-axis.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -YAxisTitleBold

Sets the Y-axis title in bold face.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -YAxisTitleSize

Sets the font size for the Y-axis title

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -YAxisNumberformat

A number formatting string, like "\#,\#\#0.00", for numbers on the Y-axis.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -YMajorUnit

Spacing for the major gridlines / tick marks on the Y-axis.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -YMinorUnit

Spacing for the minor gridlines / tick marks on the Y-axis.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -YMaxValue

Maximum value on the Y-axis.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -YMinValue

Minimum value on the Y-axis.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -YAxisPosition

Position for the Y-axis \(Left or Right\).

```yaml
Type: eAxisPosition
Parameter Sets: (All)
Aliases:
Accepted values: Left, Bottom, Right, Top

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru

Add-Excel chart doesn't normally return anything, but if -PassThru is specified it returns the newly created chart to allow it to be fine tuned.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about\_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### OfficeOpenXml.Drawing.Chart.ExcelChart

## NOTES

## RELATED LINKS

