---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Expand-NumberFormat

## SYNOPSIS

Converts short names for number formats to the formatting strings used in Excel

## SYNTAX

```text
Expand-NumberFormat [[-NumberFormat] <Object>] [<CommonParameters>]
```

## DESCRIPTION

Where you can type a number format you can write, for example, 'Short-Date' and the module will translate it into the format string used by Excel. Some formats, like Short-Date, change when Excel loads \(so date will use the local ordering of year, month and Day\). Excel also changes how markers in the are presented different cultures "," is used in the format string to mean "local thousand seperator" but depending on the country "," or "." or " " may used as the thousand seperator.

## EXAMPLES

### EXAMPLE 1

```text
Expand-NumberFormat percentage
```

Returns "0.00%"

### EXAMPLE 2

```text
Expand-NumberFormat Currency
```

Returns the currency format specified in the local regional settings, which may not be the same as Excel uses.

The regional settings set the currency symbol and then whether it is before or after the number and separated with a space or not; for negative numbers the number may be wrapped in parentheses or a - sign might appear before or after the number and symbol.

So this returns $\#,\#\#0.00;\($\#,\#\#0.00\) for English US, \#,\#\#0.00 €;€\#,\#\#0.00- for French.

Note some Eurozone countries write €1,23 and others 1,23€. In French the decimal point will be rendered as a "," and the thousand separator as a space.

## PARAMETERS

### -NumberFormat

The format string to Expand

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about\_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### System.String

## NOTES

## RELATED LINKS

