---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: 'https://github.com/dfinke/ImportExcel'
schema: 2.0.0
---

# Join-Worksheet

## SYNOPSIS

Combines data on all the sheets in an Excel worksheet onto a single sheet.

## SYNTAX

### Default \(Default\)

```text
Join-Worksheet [[-Path] <String>] [-WorkSheetName <Object>] [-Clearsheet] [-NoHeader] [-FromLabel <String>] [-LabelBlocks] [-AutoSize] [-FreezeTopRow] [-FreezeFirstColumn] [-FreezeTopRowFirstColumn] [-FreezePane <Int32[]>] [-AutoFilter] [-BoldTopRow] [-HideSource] [-Title <String>] [-TitleFillPattern <ExcelFillStyle>] [-TitleBackgroundColor <Object>] [-TitleBold] [-TitleSize <Int32>] [-PivotTableDefinition <Hashtable>] [-ExcelChartDefinition <Object[]>] [-ConditionalFormat <Object[]>] [-ConditionalText <Object[]>] [-AutoNameRange] [-RangeName <String>] [-ReturnRange] [-Show] [-PassThru][<CommonParameters>]
```

### Table

```text
Join-Worksheet [[-Path] <String>] [-WorkSheetName <Object>] [-Clearsheet] [-NoHeader] [-FromLabel <String>] [-LabelBlocks] [-AutoSize] [-FreezeTopRow] [-FreezeFirstColumn] [-FreezeTopRowFirstColumn] [-FreezePane <Int32[]>] [-BoldTopRow] [-HideSource] [-Title <String>] [-TitleFillPattern <ExcelFillStyle>] [-TitleBackgroundColor <Object>] [-TitleBold] [-TitleSize <Int32>] [-PivotTableDefinition <Hashtable>] [-ExcelChartDefinition <Object[]>] [-ConditionalFormat <Object[]>] [-ConditionalText <Object[]>] [-AutoNameRange] [-RangeName <String>] -TableName <String> [-TableStyle <TableStyles>] [-ReturnRange] [-Show] [-PassThru] [<CommonParameters>]
```

### PackageTable

```text
Join-Worksheet -ExcelPackage <ExcelPackage> [-WorkSheetName <Object>] [-Clearsheet] [-NoHeader] [-FromLabel <String>] [-LabelBlocks] [-AutoSize] [-FreezeTopRow] [-FreezeFirstColumn] [-FreezeTopRowFirstColumn] [-FreezePane <Int32[]>] [-BoldTopRow] [-HideSource] [-Title <String>] [-TitleFillPattern <ExcelFillStyle>] [-TitleBackgroundColor <Object>] [-TitleBold] [-TitleSize <Int32>] [-PivotTableDefinition <Hashtable>] [-ExcelChartDefinition <Object[]>] [-ConditionalFormat <Object[]>] [-ConditionalText <Object[]>] [-AutoNameRange] [-RangeName <String>] -TableName <String> [-TableStyle <TableStyles>] [-ReturnRange] [-Show] [-PassThru] [<CommonParameters>]
```

### PackageDefault

```text
Join-Worksheet -ExcelPackage <ExcelPackage> [-WorkSheetName <Object>] [-Clearsheet] [-NoHeader] [-FromLabel <String>] [-LabelBlocks] [-AutoSize] [-FreezeTopRow] [-FreezeFirstColumn] [-FreezeTopRowFirstColumn] [-FreezePane <Int32[]>] [-AutoFilter] [-BoldTopRow] [-HideSource] [-Title <String>] [-TitleFillPattern <ExcelFillStyle>] [-TitleBackgroundColor <Object>] [-TitleBold] [-TitleSize <Int32>] [-PivotTableDefinition <Hashtable>] [-ExcelChartDefinition <Object[]>] [-ConditionalFormat <Object[]>] [-ConditionalText <Object[]>] [-AutoNameRange] [-RangeName <String>] [-ReturnRange] [-Show] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION

Join-Worksheet can work in two main ways, either

_Combining data which has the same layout from many pages into one, or_ Combining pages which have nothing in common. In the former case the header row is copied from the first sheet and, by default, each row of data is labelled with the name of the sheet it came from.

In the latter case -NoHeader is specified, and each copied block can have the sheet it came from placed above it as a title.

## EXAMPLES

### EXAMPLE 1

```text
PS\> foreach ($computerName in @('Server1', 'Server2', 'Server3', 'Server4')) {
    Get-Service -ComputerName $computerName |
        Select-Object -Property Status, Name, DisplayName, StartType |
            Export-Excel -Path .\test.xlsx -WorkSheetname $computerName -AutoSize
}
PS\> $ptDef = New-PivotTableDefinition -PivotTableName "Pivot1" -SourceWorkSheet "Combined" -PivotRows "Status" -PivotFilter "MachineName" -PivotData @{Status='Count'} -IncludePivotChart -ChartType BarClustered3D
PS\> Join-Worksheet -Path .\test.xlsx -WorkSheetName combined -FromLabel "MachineName" -HideSource-AutoSize -FreezeTopRow -BoldTopRow -PivotTableDefinition $pt -Show
```

The foreach command gets a list of services running on four servers and exports each list to its own page in Test.xlsx. And $PtDef=... creates a definition for a PivotTable.

The Join-Worksheet command uses the same file and merges the results into a sheet named "Combined". It sets a column header of "Machinename", this column will contain the name of the sheet the data was copied from; after copying the data to the sheet "Combined", the other sheets will be hidden. Join-Worksheet finishes by calling Export-Excel to AutoSize cells, freeze the top row and make it bold and add thePivotTable.

### EXAMPLE 2

```text
PS\> Get-CimInstance -ClassName win32_logicaldisk |
        Select-Object -Property DeviceId,VolumeName, Size,Freespace |
            Export-Excel -Path "$env:computerName.xlsx" -WorkSheetname Volumes -NumberFormat "0,000"
PS\> Get-NetAdapter| Select-Object Name,InterfaceDescription,MacAddress,LinkSpeed |
          Export-Excel -Path "$env:COMPUTERNAME.xlsx" -WorkSheetname NetAdapter
PS\> Join-Worksheet -Path "$env:COMPUTERNAME.xlsx"-WorkSheetName Summary -Title "Summary" -TitleBold -TitleSize 22 -NoHeader -LabelBlocks -AutoSize -HideSource -show
```

The first two commands get logical-disk and network-card information; each type is exported to its own sheet in a workbook.

The Join-Worksheet command copies both onto a page named "Summary".Because the data is dissimilar, -NoHeader is specified, ensuring the whole of each page is copied. Specifying -LabelBlocks causes each sheet's name to become a title on the summary page above the copied data. The source data is hidden, a title is added in 22 point boldface and the columns are sized to fit the data.

## PARAMETERS

### -Path

Path to a new or existing .XLSX file.

```yaml
Type: String
Parameter Sets: Default, Table
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcelPackage

An object representing an Excel Package - either from Open-ExcelPackage or specifying -PassThru to Export-Excel.

```yaml
Type: ExcelPackage
Parameter Sets: PackageTable, PackageDefault
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WorkSheetName

The name of a sheet within the workbook where the other sheets will be joined together - "Combined" by default.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Combined
Accept pipeline input: False
Accept wildcard characters: False
```

### -Clearsheet

If specified ,any pre-existing target for the joined data will be deleted and re-created; otherwise data will be appended on this sheet.

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

### -NoHeader

Join-Worksheet assumes each sheet has identical headers and the headers should be copied to the target sheet, unless -NoHeader is specified.

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

### -FromLabel

If -NoHeader is NOT specified, then rows of data will be labeled with the name of the sheet they came from. FromLabel is the header for this column. If it is null or empty, the labels will be omitted.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: From
Accept pipeline input: False
Accept wildcard characters: False
```

### -LabelBlocks

If specified, the copied blocks of data will have the name of the sheet they were copied from inserted above them as a title.

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

### -AutoSize

Sets the width of the Excel columns to display all the data in their cells.

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

### -FreezeTopRow

Freezes headers etc. in the top row.

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

### -FreezeFirstColumn

Freezes titles etc. in the left column.

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

### -FreezeTopRowFirstColumn

Freezes top row and left column \(equivalent to Freeze pane 2,2 \).

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

### -FreezePane

Freezes panes at specified coordinates \(in the formRowNumber , ColumnNumber\).

```yaml
Type: Int32[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoFilter

Enables the Excel filter on the headers of the combined sheet.

```yaml
Type: SwitchParameter
Parameter Sets: Default, PackageDefault
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -BoldTopRow

Makes the top row boldface.

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

### -HideSource

If specified, hides the sheets that the data is copied from.

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

### -Title

Text of a title to be placed in Cell A1.

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

### -TitleFillPattern

Sets the fill pattern for the title cell.

```yaml
Type: ExcelFillStyle
Parameter Sets: (All)
Aliases:
Accepted values: None, Solid, DarkGray, MediumGray, LightGray, Gray125, Gray0625, DarkVertical, DarkHorizontal, DarkDown, DarkUp, DarkGrid, DarkTrellis, LightVertical, LightHorizontal, LightDown, LightUp, LightGrid, LightTrellis

Required: False
Position: Named
Default value: Solid
Accept pipeline input: False
Accept wildcard characters: False
```

### -TitleBackgroundColor

Sets the cell background color for the title cell.

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

### -TitleBold

Sets the title in boldface type.

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
Default value: 22
Accept pipeline input: False
Accept wildcard characters: False
```

### -PivotTableDefinition

Hashtable\(s\) with Sheet PivotRows, PivotColumns, PivotData, IncludePivotChart and ChartType values to specify a definition for one or morePivotTable\(s\).

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

### -ExcelChartDefinition

A hashtable containing ChartType, Title, NoLegend, ShowCategory, ShowPercent, Yrange, Xrange and SeriesHeader for one or more \[non-pivot\] charts.

```yaml
Type: Object[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConditionalFormat

One or more conditional formatting rules defined with New-ConditionalFormattingIconSet.

```yaml
Type: Object[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConditionalText

Applies a Conditional formatting rule defined with New-ConditionalText.

```yaml
Type: Object[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoNameRange

Makes each column a named range.

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

### -RangeName

Makes the data in the worksheet a named range.

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

### -TableName

Makes the data in the worksheet a table with a name and applies a style to it. Name must not contain spaces.

```yaml
Type: String
Parameter Sets: Table, PackageTable
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableStyle

Selects the style for the named table - defaults to "Medium6".

```yaml
Type: TableStyles
Parameter Sets: Table, PackageTable
Aliases:
Accepted values: None, Custom, Light1, Light2, Light3, Light4, Light5, Light6, Light7, Light8, Light9, Light10, Light11, Light12, Light13, Light14, Light15, Light16, Light17, Light18, Light19, Light20, Light21, Medium1, Medium2, Medium3, Medium4, Medium5, Medium6, Medium7, Medium8, Medium9, Medium10, Medium11, Medium12, Medium13, Medium14, Medium15, Medium16, Medium17, Medium18, Medium19, Medium20, Medium21, Medium22, Medium23, Medium24, Medium25, Medium26, Medium27, Medium28, Dark1, Dark2, Dark3, Dark4, Dark5, Dark6, Dark7, Dark8, Dark9, Dark10, Dark11

Required: False
Position: Named
Default value: Medium6
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReturnRange

If specified, returns the range of cells in the combined sheet, in the format "A1:Z100".

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

### -Show

Opens the Excel file immediately after creation. Convenient for viewing the results instantly without having to search for the file first.

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

If specified, an object representing the unsaved Excel package will be returned, it then needs to be saved.

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

