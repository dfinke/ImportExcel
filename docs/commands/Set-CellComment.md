---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: https://dfinke.github.io/ImportExcel/commands/Set-CellComment
schema: 2.0.0
---

# Set-CellComment

## SYNOPSIS
{{ Fill in the Synopsis }}

## SYNTAX

### Range (Default)
```
Set-CellComment [-Worksheet <ExcelWorksheet>] [-Range] <Object> -Text <String> [<CommonParameters>]
```

### ColumnNumber
```
Set-CellComment -Worksheet <ExcelWorksheet> -Row <Int32> -ColumnNumber <Int32> -Text <String>
 [<CommonParameters>]
```

### ColumnLetter
```
Set-CellComment -Worksheet <ExcelWorksheet> -Row <Int32> -ColumnLetter <String> -Text <String>
 [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1

```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -ColumnLetter
{{ Fill ColumnLetter Description }}

```yaml
Type: String
Parameter Sets: ColumnLetter
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnNumber
{{ Fill ColumnNumber Description }}

```yaml
Type: Int32
Parameter Sets: ColumnNumber
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range
{{ Fill Range Description }}

```yaml
Type: Object
Parameter Sets: Range
Aliases: Address

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Row
{{ Fill Row Description }}

```yaml
Type: Int32
Parameter Sets: ColumnNumber, ColumnLetter
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
{{ Fill Text Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Worksheet
{{ Fill Worksheet Description }}

```yaml
Type: ExcelWorksheet
Parameter Sets: Range
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

```yaml
Type: ExcelWorksheet
Parameter Sets: ColumnNumber, ColumnLetter
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.Object

## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS
