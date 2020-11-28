---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: 'https://github.com/dfinke/ImportExcel'
schema: 2.0.0
---

# Get-ExcelSheetInfo

## SYNOPSIS

Get worksheet names and their indices of an Excel workbook.

## SYNTAX

```text
Get-ExcelSheetInfo [-Path] <Object> [<CommonParameters>]
```

## DESCRIPTION

The Get-ExcelSheetInfo cmdlet gets worksheet names and their indices of an Excel workbook.

## EXAMPLES

### EXAMPLE 1

```text
Get-ExcelSheetInfo .\Test.xlsx
```

## PARAMETERS

### -Path

Specifies the path to the Excel file. \(This parameter is required.\)

```yaml
Type: Object
Parameter Sets: (All)
Aliases: FullName

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about\_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

CHANGELOG 2016/01/07 Added Created by Johan Akerstrom \([https://github.com/CosmosKey](https://github.com/CosmosKey)\)

## RELATED LINKS

[https://github.com/dfinke/ImportExcel](https://github.com/dfinke/ImportExcel)

