---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Add-PivotTable

## SYNOPSIS

Adds a PivotTable \(and optional PivotChart\) to a workbook.

## SYNTAX

### ChartbyParams \(Default\)

```text
Add-PivotTable -PivotTableName <String> [-Address <ExcelAddressBase>] [-ExcelPackage <Object>]  [-SourceWorkSheet <Object>] [-SourceRange <Object>] [-PivotRows <Object>] [-PivotData <Object>]  [-PivotColumns <Object>] [-PivotFilter <Object>] [-PivotDataToColumn] [-PivotTotals <String>] [-NoTotalsInPivot] [-GroupDateRow <String>] [-GroupDatePart <eDateGroupBy[]>] [-GroupNumericRow <String>] [-GroupNumericMin <Double>] [-GroupNumericMax <Double>] [-GroupNumericInterval <Double>] [-PivotNumberFormat <String>] [-PivotTableStyle <TableStyles>] [-IncludePivotChart] [-ChartTitle <String>]  [-ChartHeight <Int32>] [-ChartWidth <Int32>] [-ChartRow <Int32>] [-ChartColumn <Int32>] [-ChartRowOffSetPixels <Int32>] [-ChartColumnOffSetPixels <Int32>] [-ChartType <eChartType>] [-NoLegend] [-ShowCategory] [-ShowPercent] [-Activate] [-PassThru] [<CommonParameters>]
```

### ChartbyDef

```text
Add-PivotTable -PivotTableName <String> [-Address <ExcelAddressBase>] [-ExcelPackage <Object>]  [-SourceWorkSheet <Object>] [-SourceRange <Object>] [-PivotRows <Object>] [-PivotData <Object>] [-PivotColumns <Object>] [-PivotFilter <Object>] [-PivotDataToColumn] [-PivotTotals <String>] [-NoTotalsInPivot] [-GroupDateRow <String>] [-GroupDatePart <eDateGroupBy[]>] [-GroupNumericRow <String>] [-GroupNumericMin <Double>] [-GroupNumericMax <Double>] [-GroupNumericInterval <Double>] [-PivotNumberFormat <String>] [-PivotTableStyle <TableStyles>] -PivotChartDefinition <Object> [-Activate] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION

If the PivotTable already exists, the source data will be updated.

## EXAMPLES

### EXAMPLE 1

```text
PS\> $excel = Get-Service | Export-Excel -Path test.xlsx -WorksheetName Services -PassThru -AutoSize -DisplayPropertySet -TableName ServiceTable -Title "Services on $Env:COMPUTERNAME"
    Add-PivotTable -ExcelPackage $excel  -PivotTableName ServiceSummary   -SourceRange $excel.Workbook.Worksheets\[1\].Tables\[0\].Address -PivotRows Status -PivotData Name -NoTotalsInPivot -Activate
     Close-ExcelPackage $excel -Show
```

This exports data to new workbook and creates a table with the data in it.

The Pivot table is added on its own page, the table created in the first command is used as the source for the PivotTable; which counts the service names in for each Status.

At the end the Pivot page is made active.

### EXAMPLE 2

```text
PS\>  $chartdef = New-ExcelChartDefinition -Title "Gross and net by city and product"  -ChartType ColumnClustered `
         -Column 11 -Width 500 -Height 360 -YMajorUnit 500 -YMinorUnit 100 -YAxisNumberformat "$#,##0" -LegendPosition Bottom
      $excel = ConvertFrom-Csv    @"
Product, City, Gross, Net
Apple, London , 300, 250
Orange, London , 400, 350
Banana, London , 300, 200
Orange, Paris,   600, 500
Banana, Paris,   300, 200
Apple, New York, 1200,700
"@  | Export-Excel  -Path "test.xlsx"  -TableStyle Medium13 -tablename "RawData" -PassThru
      Add-PivotTable -PivotTableName Sales -Address $excel.Workbook.Worksheets[1].Cells["F1"] `
            -SourceWorkSheet $excel.Workbook.Worksheets[1]  -PivotRows City -PivotColumns Product -PivotData @{Gross="Sum";Net="Sum"} `
            -PivotNumberFormat "$#,##0.00"  -PivotTotals Both  -PivotTableStyle Medium12 -PivotChartDefinition $chartdef
      Close-ExcelPackage -show $excel
```

This script starts by defining a chart.

Then it exports some data to an XLSX file and keeps the file open.

The next step is to add the pivot table, normally this would be on its own sheet in the workbook, but here -Address is specified to place it beside the data.

The Add-Pivot table is given the chart definition and told to create a tale using the City field to create rows, the Product field to create columns and the data should be the sum of the gross field and the sum of the net field; grand totals for both gross and net are included for rows \(Cities\) and columns \(Products\) and the data is explicitly formatted as a currency.

Note that in the chart definition the number format for the axis does not include any fraction part.

### EXAMPLE 3

```text
PS> $excel = Convertfrom-csv @"
Location,OrderDate,quantity
Boston,1/1/2017,100
New York,1/21/2017,200
Boston,1/11/2017,300
New York,1/9/2017,400
Boston,1/18/2017,500
Boston,2/1/2017,600
New York,2/21/2017,700
New York,2/11/2017,800
Boston,2/9/2017,900
Boston,2/18/2017,1000
New York,1/1/2018,100
Boston,1/21/2018,200
New York,1/11/2018,300
Boston,1/9/2018,400
New York,1/18/2018,500
Boston,2/1/2018,600
Boston,2/21/2018,700
New York,2/11/2018,800
New York,2/9/2018,900
Boston,2/18/2018,1000
"@ |  Select-Object -Property @{n="OrderDate";e={[datetime]::ParseExact($_.OrderDate,"M/d/yyyy",(Get-Culture))}},
        Location, Quantity |  Export-Excel "test2.xlsx" -PassThru -AutoSize
      Set-ExcelColumn -Worksheet $excel.sheet1 -Column 1 -NumberFormat 'Short Date'
      $pt = Add-PivotTable -PassThru -PivotTableName "ByDate" -Address $excel.Sheet1.cells["F1"] -SourceWorkSheet $excel.Sheet1  -PivotRows location,orderdate -PivotData @{'quantity'='sum'}  -GroupDateRow orderdate -GroupDatePart 'Months,Years' -PivotTotals None
      $pt.RowFields[0].SubtotalTop=$false
      $pt.RowFields[0].Compact=$false
      Close-ExcelPackage $excel -Show
```

Here the data contains dates formatted as strings using US format.

These are converted to DateTime objects before being exported to Excel; the "OrderDate" column is formatted with the local short-date style.

Then the PivotTable is added; it groups information by date and location, the date is split into years and then months.

No grand totals are displayed.

The Pivot table object is caught in a variable, and the "Location" column has its subtotal moved from the top to the bottom of each location section, and the "Compact" option is disabled to prevent "Year" moving into the same column as location.

Finally the workbook is saved and shown in Excel.

## PARAMETERS

### -PivotTableName

Name for the new PivotTable - this will be the name of a sheet in the Workbook.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Address

By default, a PivotTable will be created on its own sheet, but it can be created on an existing sheet by giving the address where the top left corner of the table should go. \(Allow two rows for the filter if one is used.\)

```yaml
Type: ExcelAddressBase
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcelPackage

An Excel-package object for the workbook.

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

### -SourceWorkSheet

Worksheet where the data is found.

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

Address range in the worksheet e.g "A10:F20" - the first row must be column names: if not specified the whole sheet will be used.

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

Fields to set as rows in the PivotTable.

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

A hash table in form "FieldName"="Function", where function is one of Average, Count, CountNums, Max, Min, Product, None, StdDev, StdDevP, Sum, Var, VarP.

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

### -PivotColumns

Fields to set as columns in the PivotTable.

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

Fields to use to filter in the PivotTable.

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

If there are multiple data items in a PivotTable, by default they are shown on separate rows; this switch makes them separate columns.

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

Define whether totals should be added to rows, columns neither, or both \(the default is both\).

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

Included for compatibility - equivalent to -PivotTotals "None".

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

The name of a row field which should be grouped by parts of the date/time \(ignored if GroupDateRow is not specified\)

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

The name of a Column field which should be grouped by parts of the date/time \(ignored if GroupDateRow is not specified\)

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

The name of a row field which should be grouped by Number \(e.g. 0-99, 100-199, 200-299 \)

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

The name of a Column field which should be grouped by Number \(e.g. 0-99, 100-199, 200-299 \)

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

Number format to apply to the data cells in the PivotTable.

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

Apply a table style to the PivotTable.

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

Use a chart definition instead of specifying chart settings one by one.

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

If specified, a chart will be included.

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

Height of the chart in Pixels \(400 by default\).

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

Width of the chart in Pixels \(600 by default\).

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

Cell position of the top left corner of the chart, there will be this number of cells to the left of the chart \(default is 4, chart starts at left edge of column E\).

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

Type of chart; defaults to "Pie".

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

If specified hides the chart legend.

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

If specified attaches the category to slices in a pie chart : not supported on all chart types, this may give errors if applied to an unsupported type.

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

If there is already content in the workbook the sheet with the PivotTable will not be active UNLESS Activate is specified.

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

### -PassThru

Return the PivotTable so it can be customized.

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

### OfficeOpenXml.Table.PivotTable.ExcelPivotTable

## NOTES

## RELATED LINKS

