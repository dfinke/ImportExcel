---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: https://dfinke.github.io/ImportExcel/commands/Read-Clipboard
schema: 2.0.0
---

# Read-Clipboard

## SYNOPSIS
Read text from clipboard and pass to either ConvertFrom-Csv or ConvertFrom-Json.
Check out the how to video - https://youtu.be/dv2GOH5sbpA

## SYNTAX

```
Read-Clipboard [[-Delimiter] <Object>] [[-Header] <Object>] [<CommonParameters>]
```

## DESCRIPTION
Read text from clipboard.
It can read CSV or JSON.
Plus, you can specify the delimiter and headers.

## EXAMPLES

### EXAMPLE 1
```
Read-Clipboard # Detects if the clipboard contains CSV, JSON, or Tab delimited data.
```

### EXAMPLE 2
```
Read-Clipboard -Delimiter '|' # Converts data using a pipe delimiter
```

### EXAMPLE 3
```
Read-Clipboard -Header 'P1', 'P2', 'P3' # Specify the header columns to be used
```

## PARAMETERS

### -Delimiter
{{ Fill Delimiter Description }}

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

### -Header
{{ Fill Header Description }}

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
