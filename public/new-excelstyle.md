---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# New-ExcelStyle

## SYNOPSIS

## SYNTAX

```text
New-ExcelStyle [[-Range] <Object>] [[-NumberFormat] <Object>] [[-BorderAround] <ExcelBorderStyle>]
 [[-BorderColor] <Object>] [[-BorderBottom] <ExcelBorderStyle>] [[-BorderTop] <ExcelBorderStyle>]
 [[-BorderLeft] <ExcelBorderStyle>] [[-BorderRight] <ExcelBorderStyle>] [[-FontColor] <Object>]
 [[-Value] <Object>] [[-Formula] <Object>] [-ArrayFormula] [-ResetFont] [-Bold] [-Italic] [-Underline]
 [[-UnderLineType] <ExcelUnderLineType>] [-StrikeThru] [[-FontShift] <ExcelVerticalAlignmentFont>]
 [[-FontName] <String>] [[-FontSize] <Single>] [[-BackgroundColor] <Object>]
 [[-BackgroundPattern] <ExcelFillStyle>] [[-PatternColor] <Object>] [-WrapText]
 [[-HorizontalAlignment] <ExcelHorizontalAlignment>] [[-VerticalAlignment] <ExcelVerticalAlignment>]
 [[-TextRotation] <Int32>] [-AutoSize] [[-Width] <Single>] [[-Height] <Single>] [-Hidden] [-Locked] [-Merge]
```

## DESCRIPTION

## EXAMPLES

### Example 1

```text
PS C:\> {{ Add example code here }}
```

## PARAMETERS

### -ArrayFormula

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoSize

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: AutoFit

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundColor

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

### -BackgroundPattern

```yaml
Type: ExcelFillStyle
Parameter Sets: (All)
Aliases:
Accepted values: None, Solid, DarkGray, MediumGray, LightGray, Gray125, Gray0625, DarkVertical, DarkHorizontal, DarkDown, DarkUp, DarkGrid, DarkTrellis, LightVertical, LightHorizontal, LightDown, LightUp, LightGrid, LightTrellis

Required: False
Position: 16
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BorderAround

```yaml
Type: ExcelBorderStyle
Parameter Sets: (All)
Aliases:
Accepted values: None, Hair, Dotted, DashDot, Thin, DashDotDot, Dashed, MediumDashDotDot, MediumDashed, MediumDashDot, Thick, Medium, Double

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BorderBottom

```yaml
Type: ExcelBorderStyle
Parameter Sets: (All)
Aliases:
Accepted values: None, Hair, Dotted, DashDot, Thin, DashDotDot, Dashed, MediumDashDotDot, MediumDashed, MediumDashDot, Thick, Medium, Double

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BorderColor

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BorderLeft

```yaml
Type: ExcelBorderStyle
Parameter Sets: (All)
Aliases:
Accepted values: None, Hair, Dotted, DashDot, Thin, DashDotDot, Dashed, MediumDashDotDot, MediumDashed, MediumDashDot, Thick, Medium, Double

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BorderRight

```yaml
Type: ExcelBorderStyle
Parameter Sets: (All)
Aliases:
Accepted values: None, Hair, Dotted, DashDot, Thin, DashDotDot, Dashed, MediumDashDotDot, MediumDashed, MediumDashDot, Thick, Medium, Double

Required: False
Position: 7
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BorderTop

```yaml
Type: ExcelBorderStyle
Parameter Sets: (All)
Aliases:
Accepted values: None, Hair, Dotted, DashDot, Thin, DashDotDot, Dashed, MediumDashDotDot, MediumDashed, MediumDashDot, Thick, Medium, Double

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontColor

```yaml
Type: Object
Parameter Sets: (All)
Aliases: ForegroundColor

Required: False
Position: 8
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontName

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 13
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontShift

```yaml
Type: ExcelVerticalAlignmentFont
Parameter Sets: (All)
Aliases:
Accepted values: None, Baseline, Subscript, Superscript

Required: False
Position: 12
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontSize

```yaml
Type: Single
Parameter Sets: (All)
Aliases:

Required: False
Position: 14
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Formula

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 10
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height

```yaml
Type: Single
Parameter Sets: (All)
Aliases:

Required: False
Position: 22
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Hidden

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HorizontalAlignment

```yaml
Type: ExcelHorizontalAlignment
Parameter Sets: (All)
Aliases:
Accepted values: General, Left, Center, CenterContinuous, Right, Fill, Distributed, Justify

Required: False
Position: 18
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Locked

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Merge

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NumberFormat

```yaml
Type: Object
Parameter Sets: (All)
Aliases: NFormat

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PatternColor

```yaml
Type: Object
Parameter Sets: (All)
Aliases: PatternColour

Required: False
Position: 17
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Address

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ResetFont

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StrikeThru

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextRotation

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 20
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UnderLineType

```yaml
Type: ExcelUnderLineType
Parameter Sets: (All)
Aliases:
Accepted values: None, Single, Double, SingleAccounting, DoubleAccounting

Required: False
Position: 11
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 9
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -VerticalAlignment

```yaml
Type: ExcelVerticalAlignment
Parameter Sets: (All)
Aliases:
Accepted values: Top, Center, Bottom, Distributed, Justify

Required: False
Position: 19
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width

```yaml
Type: Single
Parameter Sets: (All)
Aliases:

Required: False
Position: 21
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WrapText

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

## INPUTS

### None

## OUTPUTS

### System.Object

## NOTES

## RELATED LINKS

