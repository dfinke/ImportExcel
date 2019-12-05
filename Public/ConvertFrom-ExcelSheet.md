---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version:
schema: 2.0.0
---

# ConvertFrom-ExcelSheet

## SYNOPSIS
Reads an Excel file an converts the data to a delimited text file.

## SYNTAX

```
ConvertFrom-ExcelSheet [-Path] <String> [[-OutputPath] <String>] [[-SheetName] <String>] [[-Encoding] <String>]
 [[-Extension] <String>] [[-Delimiter] <String>] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
ConvertFrom-ExcelSheet .\TestSheets.xlsx .\data
```

Reads each sheet in TestSheets.xlsx and outputs it to the data directory as the sheet name with the extension .txt.

### EXAMPLE 2
```
ConvertFrom-ExcelSheet .\TestSheets.xlsx .\data sheet?0
```

Reads and outputs sheets like Sheet10 and Sheet20 form TestSheets.xlsx and outputs it to the data directory as the sheet name with the extension .txt.

## PARAMETERS

### -Path
{{ Fill Path Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases: FullName

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
{{ Fill OutputPath Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: .\
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetName
{{ Fill SheetName Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: *
Accept pipeline input: False
Accept wildcard characters: False
```

### -Encoding
{{ Fill Encoding Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: UTF8
Accept pipeline input: False
Accept wildcard characters: False
```

### -Extension
{{ Fill Extension Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: .csv
Accept pipeline input: False
Accept wildcard characters: False
```

### -Delimiter
{{ Fill Delimiter Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: ;
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
