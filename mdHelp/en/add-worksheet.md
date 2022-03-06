---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Add-WorkSheet

## SYNOPSIS

Adds a worksheet to an existing workbook.

## SYNTAX

### Package

```text
Add-WorkSheet [-ExcelPackage] <ExcelPackage> [-WorksheetName <String>] [-ClearSheet] [-MoveToStart] [-MoveToEnd] [-MoveBefore <Object>] [-MoveAfter <Object>] [-Activate] [-CopySource <ExcelWorksheet>] [-NoClobber] [<CommonParameters>]
```

### WorkBook

```text
Add-WorkSheet -ExcelWorkbook <ExcelWorkbook> [-WorksheetName <String>] [-ClearSheet] [-MoveToStart] [-MoveToEnd] [-MoveBefore <Object>] [-MoveAfter <Object>] [-Activate] [-CopySource <ExcelWorksheet>] [-NoClobber] [<CommonParameters>]
```

## DESCRIPTION

If the named worksheet already exists, the -Clearsheet parameter decides whether it should be deleted and a new one returned, or if not specified the existing sheet will be returned.

By default the sheet is created at the end of the work book, the -MoveXXXX switches allow the sheet to be \[re\]positioned at the start or before or after another sheet.

A new sheet will only be made the default sheet when excel opens if -Activate is specified.

## EXAMPLES

### EXAMPLE 1

```text
PS\> $WorksheetActors = $ExcelPackage | Add-WorkSheet -WorkSheetname Actors
```

$ExcelPackage holds an Excel package object \(returned by Open-ExcelPackage, or Export-Excel -passthru\). This command will add a sheet named 'Actors', or return the sheet if it exists, and the result is stored in $WorkSheetActors.

### EXAMPLE 2

```text
PS\> $WorksheetActors = Add-WorkSheet -ExcelPackage $ExcelPackage -WorkSheetname "Actors" -ClearSheet -MoveToStart
```

This time the Excel package object is passed as a parameter instead of piped.

If the 'Actors' sheet already exists it is deleted and re-created.

The new sheet will be created last in the workbook, and -MoveToStart Moves it to the start.

### EXAMPLE 3

```text
PS\> $null = Add-WorkSheet -ExcelWorkbook $wb -WorkSheetname $DestinationName -CopySource $sourceWs -Activate
```

This time a workbook is used instead of a package, and a worksheet is copied - $SourceWs is a worksheet object, which can come from the same workbook or a different one.

Here the new copy of the data is made the active sheet when the workbook is opened.

## PARAMETERS

### -ExcelPackage

An object representing an Excel Package.

```yaml
Type: ExcelPackage
Parameter Sets: Package
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ExcelWorkbook

An Excel Workbook to which the Worksheet will be added - a Package contains one Workbook, so you can use whichever fits at the time.

```yaml
Type: ExcelWorkbook
Parameter Sets: WorkBook
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WorksheetName

The name of the worksheet, 'Sheet1' by default.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ClearSheet

If the worksheet already exists, by default it will returned, unless -ClearSheet is specified in which case it will be deleted and re-created.

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

### -MoveToStart

If specified, the worksheet will be moved to the start of the workbook.

MoveToStart takes precedence over MoveToEnd, Movebefore and MoveAfter if more than one is specified.

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

### -MoveToEnd

If specified, the worksheet will be moved to the end of the workbook.

\(This is the default position for newly created sheets, but it can be used to move existing sheets.\)

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

### -MoveBefore

If specified, the worksheet will be moved before the nominated one \(which can be an index starting from 1, or a name\).

MoveBefore takes precedence over MoveAfter if both are specified.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MoveAfter

If specified, the worksheet will be moved after the nominated one \(which can be an index starting from 1, or a name or \*\).

If \* is used, the worksheet names will be examined starting with the first one, and the sheet placed after the last sheet which comes before it alphabetically.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Activate

If there is already content in the workbook the new sheet will not be active UNLESS Activate is specified.

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

### -CopySource

If worksheet is provided as a copy source the new worksheet will be a copy of it. The source can be in the same workbook, or in a different file.

```yaml
Type: ExcelWorksheet
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoClobber

Ignored but retained for backwards compatibility.

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

### OfficeOpenXml.ExcelWorksheet

## NOTES

## RELATED LINKS

