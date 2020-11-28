---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# New-ExcelChartDefinition

## SYNOPSIS

Creates a Definition of a chart which can be added using Export-Excel, or Add-PivotTable

## SYNTAX

```text
New-ExcelChartDefinition [[-Title] <Object>] [[-Header] <Object>] [[-ChartType] <eChartType>] [[-ChartTrendLine] <eTrendLine[]>] [[-XRange] <Object>] [[-YRange] <Object>] [[-Width] <Object>] [[-Height] <Object>] [[-Row] <Object>] [[-RowOffSetPixels] <Object>] [[-Column] <Object>] [[-ColumnOffSetPixels] <Object>] [[-LegendPosition] <eLegendPosition>] [[-LegendSize] <Object>] [-LegendBold] [-NoLegend] [-ShowCategory] [-ShowPercent] [[-SeriesHeader] <Object>] [-TitleBold] [[-TitleSize] <Int32>] [[-XAxisTitleText] <String>] [-XAxisTitleBold] [[-XAxisTitleSize] <Object>] [[-XAxisNumberformat] <String>] [[-XMajorUnit] <Object>] [[-XMinorUnit] <Object>] [[-XMaxValue] <Object>] [[-XMinValue] <Object>] [[-XAxisPosition] <eAxisPosition>] [[-YAxisTitleText] <String>] [-YAxisTitleBold] [[-YAxisTitleSize] <Object>] [[-YAxisNumberformat] <String>] [[-YMajorUnit] <Object>] [[-YMinorUnit] <Object>] [[-YMaxValue] <Object>] [[-YMinValue] <Object>] [[-YAxisPosition] <eAxisPosition>] [<CommonParameters>]
```

## DESCRIPTION

All the parameters which are passed to Add-ExcelChart can be added to a chart-definition object and passed to Export-Excel with the -ExcelChartDefinition parameter, or to Add-PivotTable with the -PivotChartDefinition parameter. This command sets up those definition objects.

## EXAMPLES

### EXAMPLE 1

```text
PS\> $cDef = New-ExcelChartDefinition  -ChartType line -XRange "X" -YRange "Sinx"  -Title "Graph of Sine X" -TitleBold -TitleSize 14 -Column 2 -ColumnOffSetPixels 35 -Width 800 -XAxisTitleText "Degrees" -XAxisTitleBold -XAxisTitleSize 12 -XMajorUnit 30 -XMinorUnit 10 -XMinValue 0 -XMaxValue 361  -XAxisNumberformat "000" -YMinValue -1.25 -YMaxValue 1.25 -YMajorUnit 0.25 -YAxisNumberformat "0.00" -YAxisTitleText "Sine" -YAxisTitleBold -YAxisTitleSize 12  -SeriesHeader "Sin(x)" -LegendSize 8 -legendBold  -LegendPosition Bottom
PS\> 0..360 | ForEach-Object {[pscustomobject][ordered]@{x = $_; Sinx = "=Sin(Radians(x)) "}} | Export-Excel -AutoNameRange -now -WorkSheetname SinX -ExcelChartDefinition $cDef -Show
```

This reworks an example from Add-Excel-Chart but here the chart is defined and the defintion stored in $cDef and then Export-Excel uses $cDef .

## PARAMETERS

### -Title

The title for the chart.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: Chart Title
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header

No longer used. This may be removed in future versions.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
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
Position: 3
Default value: ColumnStacked
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartTrendLine

Superimposes one of Excel's trenline types on the chart.

```yaml
Type: eTrendLine[]
Parameter Sets: (All)
Aliases:
Accepted values: Exponential, Linear, Logarithmic, MovingAvgerage, Polynomial, Power

Required: False
Position: 4
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
Position: 5
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
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width

Width of the chart in pixels. Defaults to 500.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
Default value: 500
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height

Height of the chart in pixels. Defaults to 350.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: 350
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row

Row position of the top left corner of the chart. 0 places it at the top of the sheet, 1 below row 1 and so on.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 9
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -RowOffSetPixels

Offset to position the chart by a fraction of a row.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 10
Default value: 10
Accept pipeline input: False
Accept wildcard characters: False
```

### -Column

Column position of the top left corner of the chart. 0 places it at the edge of the sheet, 1 to the right of column A and so on.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 11
Default value: 6
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnOffSetPixels

Offset to position the chart by a fraction of a column.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 12
Default value: 5
Accept pipeline input: False
Accept wildcard characters: False
```

### -LegendPosition

Location of the key, either "Left", "Right", "Top", "Bottom" or "TopRight".

```yaml
Type: eLegendPosition
Parameter Sets: (All)
Aliases:
Accepted values: Top, Left, Right, Bottom, TopRight

Required: False
Position: 13
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
Position: 14
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

If specified, turns off display of the key. If you only have one data series it may be preferable to use the title to say what the chart is.

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

Attaches a category label in charts which support this.

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

Attaches a percentage label in charts which support this.

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

Specifies explicit name\(s\) for the data series, which will appear in the legend/key

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 15
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
Position: 16
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
Position: 17
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
Position: 18
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
Position: 19
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
Position: 20
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
Position: 21
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
Position: 22
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
Position: 23
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -XAxisPosition

Position for the X-axis \("Top" or" Bottom"\).

```yaml
Type: eAxisPosition
Parameter Sets: (All)
Aliases:
Accepted values: Left, Bottom, Right, Top

Required: False
Position: 24
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
Position: 25
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

Sets the font size for the Y-axis title.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 26
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -YAxisNumberformat

A number formatting string, like "\#,\#\#0.00", for numbers on the Y-axis

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 27
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
Position: 28
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
Position: 29
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
Position: 30
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
Position: 31
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -YAxisPosition

Position for the Y-axis \("Left" or "Right"\).

```yaml
Type: eAxisPosition
Parameter Sets: (All)
Aliases:
Accepted values: Left, Bottom, Right, Top

Required: False
Position: 32
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about\_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS

