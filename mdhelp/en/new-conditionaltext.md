---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# New-ConditionalText

## SYNOPSIS

Creates an object which describes a conditional formatting rule for single valued rules.

## SYNTAX

```text
New-ConditionalText [[-Text] <Object>] [[-ConditionalTextColor] <Object>] [[-BackgroundColor] <Object>]  [[-Range] <String>] [[-PatternType] <ExcelFillStyle>] [[-ConditionalType] <Object>] [<CommonParameters>]
```

## DESCRIPTION

Some Conditional formatting rules don't apply styles to a cell \(IconSets and Databars\); some take two parameters \(Between\); some take none \(ThisWeek, ContainsErrors, AboveAverage etc\).The others take a single parameter \(Top, BottomPercent, GreaterThan, Contains etc\).

This command creates an object to describe the last two categories, which can then be passed to Export-Excel.

## EXAMPLES

### EXAMPLE 1

```text
PS\> $ct = New-ConditionalText -Text  'Ferrari'
PS\> Export-Excel -ExcelPackage $excel -ConditionalTest $ct -show
```

The first line creates a definition object which will highlight the word "Ferrari" in any cell. and the second uses Export-Excel with an open package to apply the format and save and open the file.

### EXAMPLE 2

```text
PS\> $ct  = New-ConditionalText -Text "Ferrari"
PS\> $ct2 = New-ConditionalText -Range $worksheet.Names\["FinishPosition"\].Address -ConditionalType LessThanOrEqual -Text 3 -ConditionalTextColor Red -BackgroundColor White
PS\> Export-Excel -ExcelPackage $excel -ConditionalText $ct,$ct2 -show
```

This builds on the previous example, and specifies a condition of \&lt;=3 with a format of red text on a white background; this applies to a named range "Finish Position".

The range could be written -Range "C:C" to specify a named column, or -Range "C2:C102" to specify certain cells in the column.

## PARAMETERS

### -Text

The text \(or other value\) to use in the rule. Note that Equals, GreaterThan/LessThan rules require text to wrapped in double quotes.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: ConditionValue

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConditionalTextColor

The font color for the cell - by default: "DarkRed".

```yaml
Type: Object
Parameter Sets: (All)
Aliases: ForeGroundColor

Required: False
Position: 2
Default value: [System.Drawing.Color]::DarkRed
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundColor

The fill color for the cell - by default: "LightPink".

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: [System.Drawing.Color]::LightPink
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range

The range of cells that the conditional format applies to; if none is specified the range will be apply to all the data in the sheet.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PatternType

The background pattern for the cell - by default: "Solid"

```yaml
Type: ExcelFillStyle
Parameter Sets: (All)
Aliases:
Accepted values: None, Solid, DarkGray, MediumGray, LightGray, Gray125, Gray0625, DarkVertical, DarkHorizontal, DarkDown, DarkUp, DarkGrid, DarkTrellis, LightVertical, LightHorizontal, LightDown, LightUp, LightGrid, LightTrellis

Required: False
Position: 5
Default value: Solid
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConditionalType

One of the supported rules; by default "ContainsText" is selected.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: RuleType

Required: False
Position: 6
Default value: ContainsText
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about\_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS

[Add-ConditionalFormatting](new-conditionaltext.md) [New-ConditionalFormattingIconSet](new-conditionaltext.md)

