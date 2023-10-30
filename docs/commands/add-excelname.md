---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Add-ExcelName

## SYNOPSIS

Adds a named-range to an existing Excel worksheet.

## SYNTAX

```text
Add-ExcelName [-Range] <ExcelRange> [[-RangeName] <String>] [<CommonParameters>]
```

## DESCRIPTION

It is often helpful to be able to refer to sets of cells with a name rather than using their co-ordinates; Add-ExcelName sets up these names.

## EXAMPLES

### EXAMPLE 1

```text
PS\> Add-ExcelName -Range $ws.Cells[$dataRange] -RangeName $rangeName
```

$WS is a worksheet, and $dataRange is a string describing a range of cells - for example "A1:Z10" - which will become a named range, using the name in $rangeName.

## PARAMETERS

### -Range

The range of cells to assign as a name.

```yaml
Type: ExcelRange
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RangeName

The name to assign to the range. If the name exists it will be updated to the new range. If no name is specified, the first cell in the range will be used as the name.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about\_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS

