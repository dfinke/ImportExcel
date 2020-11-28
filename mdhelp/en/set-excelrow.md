---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Set-ExcelRow

## SYNOPSIS

Fills values into a \[new\] row in an Excel spreadsheet, and sets row formats.

## SYNTAX

### Package

```text
Set-ExcelRow -ExcelPackage <ExcelPackage> [-Worksheetname <Object>] [-Row <Object>] [-StartColumn <Int32>] [-Value <Object>] [-Heading <Object>] [-HeadingBold] [-HeadingSize <Int32>] [-NumberFormat <Object>] [-BorderAround <ExcelBorderStyle>] [-BorderColor <Object>] [-BorderBottom <ExcelBorderStyle>] [-BorderTop <ExcelBorderStyle>] [-BorderLeft <ExcelBorderStyle>] [-BorderRight <ExcelBorderStyle>] [-FontColor <Object>] [-Bold] [-Italic] [-Underline] [-UnderLineType <ExcelUnderLineType>] [-StrikeThru]
 [-FontShift <ExcelVerticalAlignmentFont>] [-FontName <String>] [-FontSize <Single>] [-BackgroundColor <Object>] [-BackgroundPattern <ExcelFillStyle>] [-PatternColor <Object>] [-WrapText] [-HorizontalAlignment <ExcelHorizontalAlignment>] [-VerticalAlignment <ExcelVerticalAlignment>] [-TextRotation <Int32>] [-Height <Single>] [-Hide] [-ReturnRange] [-PassThru] [<CommonParameters>]
```

### Sheet

```text
Set-ExcelRow -Worksheet <ExcelWorksheet> [-Row <Object>] [-StartColumn <Int32>] [-Value <Object>]  [-Heading <Object>] [-HeadingBold] [-HeadingSize <Int32>] [-NumberFormat <Object>] [-BorderAround <ExcelBorderStyle>] [-BorderColor <Object>] [-BorderBottom <ExcelBorderStyle>]  [-BorderTop <ExcelBorderStyle>] [-BorderLeft <ExcelBorderStyle>] [-BorderRight <ExcelBorderStyle>] [-FontColor <Object>] [-Bold] [-Italic] [-Underline] [-UnderLineType <ExcelUnderLineType>] [-StrikeThru] [-FontShift <ExcelVerticalAlignmentFont>] [-FontName <String>] [-FontSize <Single>] [-BackgroundColor <Object>] [-BackgroundPattern <ExcelFillStyle>] [-PatternColor <Object>] [-WrapText] [-HorizontalAlignment <ExcelHorizontalAlignment>] [-VerticalAlignment <ExcelVerticalAlignment>] [-TextRotation <Int32>] [-Height <Single>] [-Hide] [-ReturnRange] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION

Set-ExcelRow accepts either a Worksheet object or an ExcelPackage object returned by Export-Excel and the name of a sheet, and inserts the chosen contents into a row of the sheet.

The contents can be a constant, like "42", a formula or a script block which is converted into a constant or a formula.

The first cell of the row can optionally be given a heading.

## EXAMPLES

### EXAMPLE 1

```text
PS\> Set-ExcelRow -Worksheet $ws -Heading Total -Value {"=sum($columnName`2:$columnName$endrow)" }
```

$Ws contains a worksheet object, and no Row number is specified so Set-ExcelRow will select the next row after the end of the data in the sheet.

The first cell in the row will contain "Total", and each of the other cells will contain =Sum\(xx2:xx99\) where xx is the column name, and 99 is the last row of data.

Note the use of the backtick in the script block \(\`2\) to Prevent 2 becoming part of the variable "ColumnName"

The script block can use $Worksheet, $Row, $Column \(number\), $ColumnName \(letter\), $StartRow/Column and $EndRow/Column.

### EXAMPLE 2

```text
PS\> Set-ExcelRow -Worksheet $ws -Heading Total -HeadingBold -Value {"=sum($columnName`2:$columnName$endrow)" } -NumberFormat 'Currency' -StartColumn 2 -Bold -BorderTop Double -BorderBottom Thin
```

This builds on the previous example, but this time the label "Total" appears in column 2 and the formula fills from column 3 onwards.

The formula and heading are set in bold face, and the formula is formatted for the local currency, and given a double line border above and single line border below.

## PARAMETERS

### -ExcelPackage

An Excel package object - for example from Export-Excel -PassThru - if specified requires a sheet name to be passed a parameter.

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

The name of the sheet to update in the package.

```yaml
Type: Object
Parameter Sets: Package
Aliases:

Required: False
Position: Named
Default value: Sheet1
Accept pipeline input: False
Accept wildcard characters: False
```

### -Worksheet

A worksheet object instead of passing a name and package.

```yaml
Type: ExcelWorksheet
Parameter Sets: Sheet
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row

Row to fill right - first row is 1. 0 will be interpreted as first unused row.

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

### -StartColumn

Position in the row to start from.

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

Value, Formula or ScriptBlock to fill in. A ScriptBlock can use $worksheet, $row, $Column \[number\], $ColumnName \[letter\(s\)\], $startRow, $startColumn, $endRow, $endColumn.

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

Optional row-heading.

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

### -HeadingBold

Set the heading in bold type.

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

### -HeadingSize

Change the font-size of the heading.

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

### -NumberFormat

Number format to apply to cells e.g. "dd/MM/yyyy HH:mm", "£\#,\#\#0.00;\[Red\]-£\#,\#\#0.00", "0.00%" , "\#\#/\#\#" , "0.0E+0" etc.

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

### -BorderColor

Color of the border.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: [System.Drawing.Color]::Black
Accept pipeline input: False
Accept wildcard characters: False
```

### -BorderBottom

Style for the bottom border.

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

### -BorderTop

Style for the top border.

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

### -BorderLeft

Style for the left border.

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

### -BorderRight

Style for the right border.

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

Color for the text - if not specified the font will be left as it it is.

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

Subscript or Superscript \(or none\).

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

Background pattern - solid by default.

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

Position cell contents to Left, Right, Center etc. default is 'General'.

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

Degrees to rotate text. Up to +90 for anti-clockwise \("upwards"\), or to -90 for clockwise.

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

### -Height

Set cells to a fixed height.

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

### -Hide

Hide the row.

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

### -ReturnRange

If sepecified, returns the range of cells which were affected.

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

If Specified, return a row object to allow further work to be done.

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

### OfficeOpenXml.ExcelRow

### System.String

## NOTES

## RELATED LINKS

