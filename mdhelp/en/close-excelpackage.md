---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Close-ExcelPackage

## SYNOPSIS

Closes an Excel Package, saving, saving under a new name or abandoning changes and opening the file in Excel as required.

## SYNTAX

```text
Close-ExcelPackage [-ExcelPackage] <ExcelPackage> [-Show] [-NoSave] [[-SaveAs] <Object>] [[-Password] <String>] [-Calculate] [<CommonParameters>]
```

## DESCRIPTION

When working with an ExcelPackage object, the Workbook is held in memory and not saved until the .Save\(\) method of the package is called.

Close-ExcelPackage saves and disposes of the Package object.

It can be called with -NoSave to abandon the file without saving, with a new "SaveAs" filename, and/or with a password to protect the file. And -Show will open the file in Excel; -Calculate will try to update the workbook, although not everything can be recalculated

## EXAMPLES

### EXAMPLE 1

```text
Close-ExcelPackage -show $excel
```

$excel holds a package object, this saves the workbook and loads it into Excel.

### EXAMPLE 2

```text
Close-ExcelPackage -NoSave $excel
```

$excel holds a package object, this disposes of it without writing it to disk.

## PARAMETERS

### -ExcelPackage

Package to close.

```yaml
Type: ExcelPackage
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Show

Open the file in Excel.

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

### -NoSave

Abandon the file without saving.

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

### -SaveAs

Save file with a new name \(ignored if -NoSave Specified\).

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

### -Password

Password to protect the file.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Calculate

Attempt to recalculation the workbook before saving

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

