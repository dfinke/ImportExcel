---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Set-ExcelRange

## SYNOPSIS

Applies number, font, alignment and/or color formatting, values or formulas to a range of Excel cells.

## SYNTAX

```text
Set-ExcelRange [[-Range] <Object>] [-WorkSheet <ExcelWorksheet>] [-NumberFormat <Object>] [-BorderAround <ExcelBorderStyle>] [-BorderColor <Object>] [-BorderBottom <ExcelBorderStyle>] [-BorderTop <ExcelBorderStyle>] [-BorderLeft <ExcelBorderStyle>] [-BorderRight <ExcelBorderStyle>] [-FontColor <Object>] [-Value <Object>] [-Formula <Object>] [-ArrayFormula] [-ResetFont] [-Bold] [-Italic] [-Underline] [-UnderLineType <ExcelUnderLineType>] [-StrikeThru] [-FontShift <ExcelVerticalAlignmentFont>] [-FontName <String>] [-FontSize <Single>] [-BackgroundColor <Object>] [-BackgroundPattern <ExcelFillStyle>] [-PatternColor <Object>] [-WrapText] [-HorizontalAlignment <ExcelHorizontalAlignment>] [-VerticalAlignment <ExcelVerticalAlignment>] [-TextRotation <Int32>] [-AutoSize] [-Width <Single>] [-Height <Single>] [-Hidden] [-Locked] [-Merge] [<CommonParameters>]
```

## DESCRIPTION

Set-ExcelRange was created to set the style elements for a range of cells, this includes auto-sizing and hiding, setting font elements \(Name, Size, Bold, Italic, Underline & UnderlineStyle and Subscript & SuperScript\), font and background colors, borders, text wrapping, rotation, alignment within cells, and number format.

It was orignally named "Set-Format", but it has been extended to set Values, Formulas and ArrayFormulas \(sometimes called Ctrl-shift-Enter \[CSE\] formulas\); because of this, the name has become Set-ExcelRange but the old name of Set-Format is preserved as an alias.

## EXAMPLES

### EXAMPLE 1

```text
PS\> $sheet.Column(3) | Set-ExcelRange -HorizontalAlignment Right -NumberFormat "#,###" -AutoFit
```

Selects column 3 from a sheet object \(within a workbook object, which is a child of the ExcelPackage object\) and passes it to Set-ExcelRange which formats numbers as a integers with comma-separated groups, aligns it right, and auto-fits the column to the contents.

### EXAMPLE 2

```text
PS\> Set-ExcelRange -Range $sheet.Cells["E1:H1048576"]  -HorizontalAlignment Right -NumberFormat "#,###"
```

Instead of piping the address, this version specifies a block of cells and applies similar formatting.

### EXAMPLE 3

```text
PS\> Set-ExcelRange $excel.Workbook.Worksheets[1].Tables["Processes"] -Italic
```

This time instead of specifying a range of cells, a table is selected by name and formatted as italic.

## PARAMETERS

### -Range

One or more row\(s\), Column\(s\) and/or block\(s\) of cells to format.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Address

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -WorkSheet

The worksheet where the format is to be applied.

```yaml
Type: ExcelWorksheet
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

Style of border to draw around the range.

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

Colour for the text - if none is specified it will be left as it is.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: ForegroundColor

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value

Value for the cell.

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

### -Formula

Formula for the cell.

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

### -ArrayFormula

Specifies formula should be an array formula \(a.k.a CSE \[ctrl-shift-enter\] formula\).

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

### -ResetFont

Clear Bold, Italic, StrikeThrough and Underline and set color to Black.

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

Strike through text; use -Strikethru:$false to remove Strike through

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

Background pattern - Solid by default.

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

Autofit cells to width \(columns or ranges only\).

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

Set cells to a fixed width \(columns or ranges only\), ignored if Autosize is specified.

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

### -Height

Set cells to a fixed height \(rows or ranges only\).

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

### -Hidden

Hide a row or column \(not a range\); use -Hidden:$false to unhide.

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

### -Locked

Locks cells. Cells are locked by default use -locked:$false on the whole sheet and then lock specific ones, and enable protection on the sheet.

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

### -Merge

Merges cells - it is recommended that you explicitly set -HorizontalAlignment

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

