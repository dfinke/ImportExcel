---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: https://dfinke.github.io/ImportExcel/commands/Get-ExcelFileSchema
schema: 2.0.0
---

# Get-ExcelFileSchema

## SYNOPSIS
Gets the schema of an Excel file.

## SYNTAX

```
Get-ExcelFileSchema [-Path] <Object> [-Compress] [<CommonParameters>]
```

## DESCRIPTION
The Get-ExcelFileSchema function gets the schema of an Excel file by returning the property names of the first row of each worksheet in the file.

## EXAMPLES

### EXAMPLE 1
```
Get-ExcelFileSchema -Path .\example.xlsx
```

## PARAMETERS

### -Compress
Indicates whether to compress the json output.

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

### -Path
Specifies the path to the Excel file.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: FullName

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### Json
## NOTES

## RELATED LINKS
