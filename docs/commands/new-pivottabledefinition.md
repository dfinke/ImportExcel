---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# New-PivotTableDefinition

## SYNOPSIS

Creates PivotTable definitons for Export-Excel

## SYNTAX

### ChartbyDef

```text
New-PivotTableDefinition -PivotTableName <Object> [-SourceWorkSheet <Object>] [-SourceRange <Object>] [-PivotRows <Object>] [-PivotData <Hashtable>] [-PivotColumns <Object>] [-PivotFilter <Object>] [-PivotDataToColumn] [-PivotTotals <String>] [-NoTotalsInPivot] [-GroupDateRow <String>] [-GroupDatePart <eDateGroupBy[]>] [-GroupNumericRow <String>] [-GroupNumericMin <Double>] [-GroupNumericMax <Double>] [-GroupNumericInterval <Double>] [-PivotNumberFormat <String>] [-PivotTableStyle <TableStyles>] -PivotChartDefinition <Object> [-Activate] [<CommonParameters>]
```

### ChartbyParams

```text
New-PivotTableDefinition -PivotTableName <Object> [-SourceWorkSheet <Object>] [-SourceRange <Object>] [-PivotRows <Object>] [-PivotData <Hashtable>] [-PivotColumns <Object>] [-PivotFilter <Object>] [-PivotDataToColumn] [-PivotTotals <String>] [-NoTotalsInPivot] [-GroupDateRow <String>] [-GroupDatePart <eDateGroupBy[]>] [-GroupNumericRow <String>] [-GroupNumericMin <Double>] [-GroupNumericMax <Double>] [-GroupNumericInterval <Double>] [-PivotNumberFormat <String>] [-PivotTableStyle <TableStyles>] [-IncludePivotChart] [-ChartTitle <String>] [-ChartHeight <Int32>] [-ChartWidth <Int32>] [-ChartRow <Int32>] [-ChartColumn <Int32>] [-ChartRowOffSetPixels <Int32>] [-ChartColumnOffSetPixels <Int32>] [-ChartType <eChartType>] [-NoLegend] [-ShowCategory] [-ShowPercent] [-Activate] [<CommonParameters>]
```

## DESCRIPTION

Export-Excel allows a single PivotTable to be defined using the parameters -IncludePivotTable, -PivotColumns, -PivotRows, -PivotData, -PivotFilter, -PivotTotals, -PivotDataToColumn, -IncludePivotChart and -ChartType.

Its -PivotTableDefintion paramater allows multiple PivotTables to be defined, with additional parameters. New-PivotTableDefinition is a convenient way to build these definitions.

## EXAMPLES

### EXAMPLE 1

```text
PS\> $pt = New-PivotTableDefinition -PivotTableName "PT1" -SourceWorkSheet "Sheet1" -PivotRows "Status" -PivotData @{Status='Count'} -PivotFilter 'StartType' -IncludePivotChart -ChartType BarClustered3D
PS\> $Pt += New-PivotTableDefinition -PivotTableName "PT2" -SourceWorkSheet "Sheet2" -PivotRows "Company" -PivotData @{Company='Count'} -IncludePivotChart -ChartType PieExploded3D -ShowPercent -ChartTitle "Breakdown of processes by company"
PS\> Get-Service | Select-Object -Property Status,Name,DisplayName,StartType | Export-Excel -Path .\test.xlsx -AutoSize
PS\> Get-Process | Select-Object -Property Name,Company,Handles,CPU,VM | Export-Excel -Path .\test.xlsx -AutoSize -WorksheetName 'sheet2'
PS\> $excel = Export-Excel -Path .\test.xlsx -PivotTableDefinition $pt -Show
```

This is a re-work of one of the examples in Export-Excel - instead of writing out the pivot definition hash-table, it is built by calling New-PivotTableDefinition.

## PARAMETERS

### -PivotTableName

Name for the new pivot table

This command previously had a typo - and has an alias to avoid breaking scripts This will be removed in a future release

```yaml
Type: Object
Parameter Sets: (All)
Aliases: PivtoTableName

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceWorkSheet

Worksheet where the data is found

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

### -SourceRange

Address range in the worksheet e.g "A10:F20" - the first row must contain the column names to pivot by: if the range is not specified the whole source sheet will be used.

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

### -PivotRows

Fields to set as rows in the PivotTable

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

### -PivotData

A hash-table in form "FieldName"="Function", where function is one of Average, Count, CountNums, Max, Min, Product, None, StdDev, StdDevP, Sum, Var, VarP

```yaml
Type: Hashtable
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PivotColumns

Fields to set as columns in the PivotTable

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

### -PivotFilter

Fields to use to filter in the PivotTable

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

### -PivotDataToColumn

If there are multiple datasets in a PivotTable, by default they are shown seperatate rows under the given row heading; this switch makes them seperate columns.

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

### -PivotTotals

By default PivotTables have Totals for each Row \(on the right\) and for each column at the bottom. This allows just one or neither to be selected.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Both
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoTotalsInPivot

Included for compatibility - equivalent to -PivotTotals "None"

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

### -GroupDateRow

The name of a row field which should be grouped by parts of the date/time \(ignored if GroupDatePart is not specified\)

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

### -GroupDateColumn

The name of a column field which should be grouped by parts of the date/time \(ignored if GroupDatePart is not specified\)

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

### -GroupDatePart

The Part\(s\) of the date to use in the grouping \(ignored if GroupDateRow is not specified\)

```yaml
Type: eDateGroupBy[]
Parameter Sets: (All)
Aliases:
Accepted values: Years, Quarters, Months, Days, Hours, Minutes, Seconds
Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GroupNumericRow

The name of a row field which should be grouped by Number \(e.g 0-99, 100-199, 200-299 \)

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


### -GroupNumericColumn

The name of a column field which should be grouped by Number \(e.g 0-99, 100-199, 200-299 \)

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


### -GroupNumericMin

The starting point for grouping

```yaml
Type: Double
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -GroupNumericMax

The endpoint for grouping

```yaml
Type: Double
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 1.79769313486232E+308
Accept pipeline input: False
Accept wildcard characters: False
```

### -GroupNumericInterval

The interval for grouping

```yaml
Type: Double
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 100
Accept pipeline input: False
Accept wildcard characters: False
```

### -PivotNumberFormat

Number format to apply to the data cells in the PivotTable

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

### -PivotTableStyle

Apply a table style to the PivotTable

```yaml
Type: TableStyles
Parameter Sets: (All)
Aliases:
Accepted values: None, Custom, Light1, Light2, Light3, Light4, Light5, Light6, Light7, Light8, Light9, Light10, Light11, Light12, Light13, Light14, Light15, Light16, Light17, Light18, Light19, Light20, Light21, Medium1, Medium2, Medium3, Medium4, Medium5, Medium6, Medium7, Medium8, Medium9, Medium10, Medium11, Medium12, Medium13, Medium14, Medium15, Medium16, Medium17, Medium18, Medium19, Medium20, Medium21, Medium22, Medium23, Medium24, Medium25, Medium26, Medium27, Medium28, Dark1, Dark2, Dark3, Dark4, Dark5, Dark6, Dark7, Dark8, Dark9, Dark10, Dark11

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PivotChartDefinition

Use a chart definition instead of specifying chart settings one by one

```yaml
Type: Object
Parameter Sets: ChartbyDef
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -IncludePivotChart

If specified a chart Will be included.

```yaml
Type: SwitchParameter
Parameter Sets: ChartbyParams
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartTitle

Optional title for the pivot chart, by default the title omitted.

```yaml
Type: String
Parameter Sets: ChartbyParams
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartHeight

Height of the chart in Pixels \(400 by default\)

```yaml
Type: Int32
Parameter Sets: ChartbyParams
Aliases:

Required: False
Position: Named
Default value: 400
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartWidth

Width of the chart in Pixels \(600 by default\)

```yaml
Type: Int32
Parameter Sets: ChartbyParams
Aliases:

Required: False
Position: Named
Default value: 600
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartRow

Cell position of the top left corner of the chart, there will be this number of rows above the top edge of the chart \(default is 0, chart starts at top edge of row 1\).

```yaml
Type: Int32
Parameter Sets: ChartbyParams
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartColumn

Cell position of the top left corner of the chart, there will be this number of cells to the left of the chart \(default is 4, chart starts at left edge of column E\)

```yaml
Type: Int32
Parameter Sets: ChartbyParams
Aliases:

Required: False
Position: Named
Default value: 4
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartRowOffSetPixels

Vertical offset of the chart from the cell corner.

```yaml
Type: Int32
Parameter Sets: ChartbyParams
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartColumnOffSetPixels

Horizontal offset of the chart from the cell corner.

```yaml
Type: Int32
Parameter Sets: ChartbyParams
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartType

Type of chart

```yaml
Type: eChartType
Parameter Sets: ChartbyParams
Aliases:
Accepted values: Area, Line, Pie, Bubble, ColumnClustered, ColumnStacked, ColumnStacked100, ColumnClustered3D, ColumnStacked3D, ColumnStacked1003D, BarClustered, BarStacked, BarStacked100, BarClustered3D, BarStacked3D, BarStacked1003D, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, PieExploded, PieExploded3D, BarOfPie, XYScatterSmooth, XYScatterSmoothNoMarkers, XYScatterLines, XYScatterLinesNoMarkers, AreaStacked, AreaStacked100, AreaStacked3D, AreaStacked1003D, DoughnutExploded, RadarMarkers, RadarFilled, Surface, SurfaceWireframe, SurfaceTopView, SurfaceTopViewWireframe, Bubble3DEffect, StockHLC, StockOHLC, StockVHLC, StockVOHLC, CylinderColClustered, CylinderColStacked, CylinderColStacked100, CylinderBarClustered, CylinderBarStacked, CylinderBarStacked100, CylinderCol, ConeColClustered, ConeColStacked, ConeColStacked100, ConeBarClustered, ConeBarStacked, ConeBarStacked100, ConeCol, PyramidColClustered, PyramidColStacked, PyramidColStacked100, PyramidBarClustered, PyramidBarStacked, PyramidBarStacked100, PyramidCol, XYScatter, Radar, Doughnut, Pie3D, Line3D, Column3D, Area3D

Required: False
Position: Named
Default value: Pie
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoLegend

If specified hides the chart legend

```yaml
Type: SwitchParameter
Parameter Sets: ChartbyParams
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowCategory

if specified attaches the category to slices in a pie chart : not supported on all chart types, this may give errors if applied to an unsupported type.

```yaml
Type: SwitchParameter
Parameter Sets: ChartbyParams
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowPercent

If specified attaches percentages to slices in a pie chart.

```yaml
Type: SwitchParameter
Parameter Sets: ChartbyParams
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Activate

If there is already content in the workbook the sheet with the PivotTable will not be active UNLESS Activate is specified

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

## NOTES

## RELATED LINKS

