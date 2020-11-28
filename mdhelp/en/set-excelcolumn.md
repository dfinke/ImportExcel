---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Set-ExcelColumn

## SYNOPSIS

Adds or modifies a column in an Excel worksheet, filling values, setting formatting and/or creating named ranges.

## SYNTAX

### Package

```text
Set-ExcelColumn -ExcelPackage <ExcelPackage> [-Worksheetname <String>] [-Column <Object>] [-StartRow <Int32>]  [-Value <Object>] [-Heading <Object>] [-NumberFormat <Object>] [-BorderAround <ExcelBorderStyle>] [-FontColor <Object>] [-Bold] [-Italic] [-Underline] [-UnderLineType <ExcelUnderLineType>] [-StrikeThru]  [-FontShift <ExcelVerticalAlignmentFont>] [-FontName <String>] [-FontSize <Single>]  [-BackgroundColor <Object>] [-BackgroundPattern <ExcelFillStyle>] [-PatternColor <Object>] [-WrapText] [-HorizontalAlignment <ExcelHorizontalAlignment>] [-VerticalAlignment <ExcelVerticalAlignment>] [-TextRotation <Int32>] [-AutoSize] [-Width <Single>] [-AutoNameRange] [-Hide] [-Specified] [-PassThru]  [<CommonParameters>]
```

### sheet

```text
Set-ExcelColumn -Worksheet <ExcelWorksheet> [-Column <Object>] [-StartRow <Int32>] [-Value <Object>] [-Heading <Object>] [-NumberFormat <Object>] [-BorderAround <ExcelBorderStyle>] [-FontColor <Object>] [-Bold] [-Italic] [-Underline] [-UnderLineType <ExcelUnderLineType>] [-StrikeThru] [-FontShift <ExcelVerticalAlignmentFont>] [-FontName <String>] [-FontSize <Single>]  [-BackgroundColor <Object>] [-BackgroundPattern <ExcelFillStyle>] [-PatternColor <Object>] [-WrapText] [-HorizontalAlignment <ExcelHorizontalAlignment>] [-VerticalAlignment <ExcelVerticalAlignment>] [-TextRotation <Int32>] [-AutoSize] [-Width <Single>] [-AutoNameRange] [-Hide] [-Specified] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION

Set-ExcelColumn can take a value which is either a string containing a value or formula or a scriptblock which evaluates to a string, and optionally a column number and fills that value down the column.

A column heading can be specified, and the column can be made a named range.

The column can be formatted in the same operation.

## EXAMPLES

### EXAMPLE 1

```text
PS\> Set-ExcelColumn -Worksheet $ws -Column 5 -NumberFormat 'Currency'
```

$ws contains a worksheet object - and column "E" is set to use the local currency format.

Intelisense will complete the names of predefined number formats.

You can see how currency is interpreted on the local computer with the command Expand-NumberFormat currency

### EXAMPLE 2

```text
PS\> Set-ExcelColumn -Worksheet $ws -Heading "WinsToFastLaps"  -Value {"=E$row/C$row"} -Column 7 -AutoSize -AutoNameRange
```

Here, $WS already contains a worksheet which holds counts of races won and fastest laps recorded by racing drivers \(in columns C and E\). Set-ExcelColumn specifies that Column 7 should have a heading of "WinsToFastLaps" and the data cells should contain =E2/C2 , =E3/C3 etc the new data cells should become a named range, which will also be named "WinsToFastLaps" and the column width will be set automatically.

When a value begins with "=", it is treated as a formula.

If value is a script block it will be evaluated, so here the string "=E$row/C$Row" will have the number of the current row inserted; see the value parameter for a list of variables which can be used.

Note than when evaluating an expression in a string, it is necessary to wrap it in $\(\) so $row is valid but $\($row+1\) is needed.

To preventVariables merging into other parts of the string, use the back tick "$columnName\`4" will be "G4" - without the backtick the string will look for a variable named "columnName4"

### EXAMPLE 3

```text
Set-ExcelColumn -Worksheet $ws -Heading "Link" -Value {"https://en.wikipedia.org" + $worksheet.cells["B$Row"].value  }  -AutoSize
```

In this example, the worksheet in $ws has partial links to Wikipedia pages in column B.

The -Value parameter is a script block which outputs a string beginning "https..." and ending with the value of the cell at column B in the current row.

When given a valid URI, Set-ExcelColumn makes it a hyperlink.

The column will be autosized to fit the links.

### EXAMPLE 4

```text
4..6 | Set-ExcelColumn -Worksheet $ws -AutoNameRange
```

Again $ws contains a worksheet. Here columns 4 to 6 are made into named ranges, row 1 is used for the range name and the rest of the column becomes the range.

## PARAMETERS

### -ExcelPackage

If specifying the worksheet by name, the ExcelPackage object which contains the worksheet also needs to be passed.

```yaml
Type: ExcelPackage
Parameter Sets: Package
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Worksheetname

The sheet to update can be given as a name or an Excel Worksheet object - this sets it by name.

```yaml
Type: String
Parameter Sets: Package
Aliases:

Required: False
Position: Named
Default value: Sheet1
Accept pipeline input: False
Accept wildcard characters: False
```

### -Worksheet

This passes the worksheet object instead of passing a sheet name and an Excelpackage object.

```yaml
Type: ExcelWorksheet
Parameter Sets: sheet
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Column

Column to fill down - the first column is 1. 0 will be interpreted as first empty column.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -StartRow

First row to fill data in.

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

### -Value

A value, formula or scriptblock to fill in. A script block can use $worksheet, $row, $column \[number\], $columnName \[letter\(s\)\], $startRow, $startColumn, $endRow, $endColumn.

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

### -Heading

Optional column heading.

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

### -NumberFormat

Number format to apply to cells for example "dd/MM/yyyy HH:mm", "£\#,\#\#0.00;\[Red\]-£\#,\#\#0.00", "0.00%" , "\#\#/\#\#" or "0.0E+0" etc.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: NFormat

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BorderAround

Style of border to draw around the row.

```yaml
Type: ExcelBorderStyle
Parameter Sets: (All)
Aliases:
Accepted values: None, Hair, Dotted, DashDot, Thin, DashDotDot, Dashed, MediumDashDotDot, MediumDashed, MediumDashDot, Thick, Medium, Double

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontColor

Colour for the text - if none specified it will be left as it it is.

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

### -Bold

Make text bold; use -Bold:$false to remove bold.

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

### -Italic

Make text italic; use -Italic:$false to remove italic.

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

### -Underline

Underline the text using the underline style in -UnderlineType; use -Underline:$false to remove underlining.

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

### -UnderLineType

Specifies whether underlining should be single or double, normal or accounting mode. The default is "Single".

```yaml
Type: ExcelUnderLineType
Parameter Sets: (All)
Aliases:
Accepted values: None, Single, Double, SingleAccounting, DoubleAccounting

Required: False
Position: Named
Default value: Single
Accept pipeline input: False
Accept wildcard characters: False
```

### -StrikeThru

Strike through text; use -StrikeThru:$false to remove strike through.

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

### -FontShift

Subscript or Superscript \(or None\).

```yaml
Type: ExcelVerticalAlignmentFont
Parameter Sets: (All)
Aliases:
Accepted values: None, Baseline, Subscript, Superscript

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontName

Font to use - Excel defaults to Calibri.

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

### -FontSize

Point size for the text.

```yaml
Type: Single
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundColor

Change background color.

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

### -BackgroundPattern

Background pattern - "Solid" by default.

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

### -PatternColor

Secondary color for background pattern.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: PatternColour

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WrapText

Turn on Text-Wrapping; use -WrapText:$false to turn off wrapping.

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

### -HorizontalAlignment

Position cell contents to Left, Right, Center etc. Default is "General".

```yaml
Type: ExcelHorizontalAlignment
Parameter Sets: (All)
Aliases:
Accepted values: General, Left, Center, CenterContinuous, Right, Fill, Distributed, Justify

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -VerticalAlignment

Position cell contents to Top, Bottom or Center.

```yaml
Type: ExcelVerticalAlignment
Parameter Sets: (All)
Aliases:
Accepted values: Top, Center, Bottom, Distributed, Justify

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextRotation

Degrees to rotate text; up to +90 for anti-clockwise \("upwards"\), or to -90 for clockwise.

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

### -AutoSize

Attempt to auto-fit cells to the width their contents.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: AutoFit

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width

Set cells to a fixed width, ignored if -AutoSize is specified.

```yaml
Type: Single
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoNameRange

Set the inserted data to be a named range.

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

### -Hide

Hide the column.

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

### -Specified

If specified, returns the range of cells which were affected.

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

If specified, return an object representing the Column, to allow further work to be done on it.

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

### OfficeOpenXml.ExcelColumn

### System.String

## NOTES

## RELATED LINKS

