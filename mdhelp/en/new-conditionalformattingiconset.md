---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# New-ConditionalFormattingIconSet

## SYNOPSIS

Creates an object which describes a conditional formatting rule a for 3,4 or 5 icon set.

## SYNTAX

```text
New-ConditionalFormattingIconSet [-Range] <Object> [[-ConditionalFormat] <Object>] [-Reverse] [<CommonParameters>]
```

## DESCRIPTION

Export-Excel takes a -ConditionalFormat parameter which can hold one or more descriptions for conditional formats; this command builds the defintion of a Conditional formatting rule for an icon set.

## EXAMPLES

### EXAMPLE 1

```text
PS\> $cfRange = [OfficeOpenXml.ExcelAddress]::new($topRow, $column, $lastDataRow, $column)
PS\> $cfdef = New-ConditionalFormattingIconSet -Range $cfrange -ConditionalFormat ThreeIconSet -IconType Arrows
PS\> Export-Excel -ExcelPackage $excel -ConditionalFormat $cfdef -show
```

The first line creates a range - one column wide in the column $column, running from $topRow to $lastDataRow. The second line creates a definition object using this range and the third uses Export-Excel with an open package to apply the format and save and open the file.

## PARAMETERS

### -Range

The range of cells that the conditional format applies to.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConditionalFormat

The type of rule: one of "ThreeIconSet","FourIconSet" or "FiveIconSet"

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

### -Reverse

Use the icons in the reverse order.

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

[Add-Add-ConditionalFormatting](new-conditionalformattingiconset.md) [New-ConditionalText](new-conditionalformattingiconset.md)

