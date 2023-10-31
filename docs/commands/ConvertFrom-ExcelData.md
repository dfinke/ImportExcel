---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: https://dfinke.github.io/ImportExcel/commands/ConvertFrom-ExcelData
schema: 2.0.0
---

# ConvertFrom-ExcelData

## SYNOPSIS
{{ Fill in the Synopsis }}

## SYNTAX

```
ConvertFrom-ExcelData [-Path] <Object> [[-ScriptBlock] <ScriptBlock>] [[-WorksheetName] <Object>]
 [[-StartRow] <Int32>] [[-Header] <String[]>] [-NoHeader] [-DataOnly] [<CommonParameters>]
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

### -DataOnly
{{ Fill DataOnly Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header
{{ Fill Header Description }}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader
{{ Fill NoHeader Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
{{ Fill Path Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases: FullName

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -ScriptBlock
{{ Fill ScriptBlock Description }}

```yaml
Type: ScriptBlock
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartRow
{{ Fill StartRow Description }}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: HeaderRow, TopRow

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WorksheetName
{{ Fill WorksheetName Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Sheet

Required: False
Position: 2
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
