---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: 'https://github.com/dfinke/ImportExcel'
schema: 2.0.0
---

# ConvertFrom-ExcelSheet

## SYNOPSIS

Exports Sheets from Excel Workbooks to CSV files.

## SYNTAX

```text
ConvertFrom-ExcelSheet [-Path] <String> [[-OutputPath] <String>] [[-SheetName] <String>] [[-Encoding] <Encoding>]
 [[-Extension] <String>] [[-Delimiter] <String>] [[-Property] <Object>] [[-ExcludeProperty] <Object>] [-Append]
 [[-AsText] <String[]>] [<CommonParameters>]
```

## DESCRIPTION

This command provides a convenient way to run Import-Excel @ImportParameters \| Select-Object @selectParameters \| export-Csv @ ExportParameters It can take the parameters -AsText , as used in Import-Excel, \)Properties & -ExcludeProperties as used in Select-Object and -Append, -Delimiter and -Encoding as used in Export-CSV

## EXAMPLES

### Example 1

```text
PS C:\> ConvertFrom-ExcelSheet Path .\__tests__\First10Races.xlsx -OutputPath .. -AsText GridPosition,date
```

First10Races.xlsx contains information about Motor races. The race date and grid \(starting\) position are stored with custom formats. The command specifies the path to the file, and the directory to create the output file, and specifies that the columns "GridPosition" and "Date" should be treated as text to preserve their formatting

### Example 2

```text
PS C:\> ConvertFrom-ExcelSheet Path .\__tests__\First10Races.xlsx -OutputPath .. -AsText "GridPosition" -Property driver, @{n="date"; e={[datetime]::FromOADate($_.Date).tostring("#MM/dd/yyyy#")}} , FinishPosition, GridPosition
```

This uses the same file as example 1. Because the race date has a custom format, it imports as a number, The requirement is to create a CSV file with the Driver, a specially formatted Date, FinishPostion and GridPostion \(keeping its custom formatting\). The command specifies the path to the file, and the directory to create the output file, specifies that the column "GridPosition" should be treated as text instead of a number, and the output properties should be Driver, a calculated "date" field, FinishPosition and GridPsition. FromOADate converts the dates used by Excel \(Days since Jan 1 1900\) to a datetime object.

## PARAMETERS

### -Append

Use this parameter to have the export add output to the end of the file. Without this parameter, the command replaces the file contents without warning.

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

### -AsText

AsText allows selected columns to be returned as the text displayed in their cells, instead of their value. \(\* is supported as a wildcard.\)

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsDate

Not all date formats are recognized as indicating the number in the cell represents a date AsDate forces the number which would be returned to be converted to a date. \(\* is supported as a wildcard.\)

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Delimiter

Selects , or ; as the delimeter for the exported data - if not specified , is used by default.

```yaml
Type: String
Parameter Sets: (All)
Aliases:
Accepted values: ;, ,

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Encoding

Sets the text encoding for the output data file; UTF8 bu default

```yaml
Type: Encoding
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeProperty

Specifies the properties that to exclude from the export. Wildcards are permitted. This parameter is effective only when the command also includes the Property parameter.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Extension

Sets the file extension for the exported data, defaults to CSV

```yaml
Type: String
Parameter Sets: (All)
Aliases:
Accepted values: .txt, .log, .csv

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath

The directory where the output file\(s\) will be created. The file name\(s\) will match the name of the workbook page which contained the data.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path

The path to the .XLSX file which will be exported.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Property

Specifies the properties to select. Wildcards are permitted - the default is "\*". The value of the Property parameter can be a new calculated property, and follows the same pattern as Select-Item

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetName

The name of a sheet to export, or a regular expression which is used to identify sheets

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about\_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### None

## OUTPUTS

### System.Object

## NOTES

## RELATED LINKS

