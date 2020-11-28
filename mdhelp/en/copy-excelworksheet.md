---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Copy-ExcelWorkSheet

## SYNOPSIS

Copies a worksheet between workbooks or within the same workbook.

## SYNTAX

```text
Copy-ExcelWorkSheet [-SourceObject] <Object> [[-SourceWorkSheet] <Object>] [-DestinationWorkbook] <Object>
 [[-DestinationWorksheet] <Object>] [-Show] [<CommonParameters>]
```

## DESCRIPTION

Copy-ExcelWorkSheet takes a Source object which is either a worksheet, or a package, Workbook or path, in which case the source worksheet can be specified by name or number \(starting from 1\). The destination worksheet can be explicitly named, or will follow the name of the source if no name is specified. The Destination workbook can be given as the path to an XLSx file, an ExcelPackage object or an ExcelWorkbook object.

## EXAMPLES

### EXAMPLE 1

```text
Copy-ExcelWorkSheet -SourceWorkbook Test1.xlsx  -DestinationWorkbook Test2.xlsx
```

This is the simplest version of the command: no source worksheet is specified so Copy-ExcelWorksheet uses the first sheet in the workbook No Destination sheet is specified so the new worksheet will be the same as the one which is being copied.

### EXAMPLE 2

```text
Copy-ExcelWorkSheet -SourceWorkbook Server1.xlsx -sourceWorksheet "Settings" -DestinationWorkbook Settings.xlsx -DestinationWorksheet "Server1"
```

Here the Settings page from Server1's workbook is copied to the 'Server1" page of a "Settings" workbook.

### EXAMPLE 3

```text
$excel = Open-ExcelPackage   .\test.xlsx
```

C:\&gt; Copy-ExcelWorkSheet -SourceWorkbook $excel -SourceWorkSheet "first" -DestinationWorkbook $excel -Show -DestinationWorksheet Duplicate This opens the workbook test.xlsx and copies the worksheet named "first" to a new worksheet named "Duplicate", because -Show is specified the file is saved and opened in Excel

### EXAMPLE 4

```text
$excel = Open-ExcelPackage .\test.xlsx
```

C:\&gt; Copy-ExcelWorkSheet -SourceWorkbook $excel -SourceWorkSheet 1 -DestinationWorkbook $excel -DestinationWorksheet Duplicate C:\&gt; Close-ExcelPackage $excel This is almost the same as the previous example, except source sheet is specified by position rather than name and because -Show is not specified, so other steps can be carried using the package object, at the end the file is saved by Close-ExcelPackage

## PARAMETERS

### -SourceObject

An ExcelWorkbook or ExcelPackage object or the path to an XLSx file where the data is found.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: SourceWorkbook

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -SourceWorkSheet

Name or number \(starting from 1\) of the worksheet in the source workbook \(defaults to 1\).

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: 1
Accept pipeline input: False
Accept wildcard characters: False
```

### -DestinationWorkbook

An ExcelWorkbook or ExcelPackage object or the path to an XLSx file where the data should be copied.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DestinationWorksheet

Name of the worksheet in the destination workbook; by default the same as the source worksheet's name. If the sheet exists it will be deleted and re-copied.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show

if the destination is an excel package or a path, launch excel and open the file on completion.

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

