---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Remove-WorkSheet

## SYNOPSIS

Removes one or more worksheets from one or more workbooks

## SYNTAX

```text
Remove-WorkSheet [[-FullName] <Object>] [[-WorksheetName] <String[]>] [-Show] [-WhatIf] [-Confirm]  [<CommonParameters>]
```

## EXAMPLES

### EXAMPLE 1

```text
PS\> Remove-WorkSheet -Path Test1.xlsx -WorksheetName Sheet1
```

Removes the worksheet named 'Sheet1' from 'Test1.xlsx'

### EXAMPLE 2

```text
PS\> Remove-WorkSheet -Path Test1.xlsx -WorksheetName Sheet1,Target1
```

Removes the worksheet named 'Sheet1' and 'Target1' from 'Test1.xlsx'

### EXAMPLE 3

```text
PS\> Remove-WorkSheet -Path Test1.xlsx -WorksheetName Sheet1,Target1 -Show
```

Removes the worksheets and then launches the xlsx in Excel

### EXAMPLE 1

```text
PS\>  dir c:\reports\*.xlsx | Remove-WorkSheet
```

Removes 'Sheet1' from all the xlsx files in the c:\reports directory

## PARAMETERS

### -FullName

The fully qualified path to the XLSX file\(s\)

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Path

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -WorksheetName

The worksheet to be removed \(sheet1 by default\)

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: Sheet1
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show

If specified the file will be opened in excel after the sheet is removed.

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

### -WhatIf

Shows what would happen if the cmdlet runs. The cmdlet is not run.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: wi

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Confirm

Prompts you for confirmation before running the cmdlet.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: cf

Required: False
Position: Named
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

