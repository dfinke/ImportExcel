---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# ConvertFrom-ExcelToSQLInsert

## SYNOPSIS

Generate SQL insert statements from Excel spreadsheet.

## SYNTAX

```text
ConvertFrom-ExcelToSQLInsert [-TableName] <Object> [-Path] <Object> [[-WorkSheetname] <Object>] [[-StartRow] <Int32>] [[-Header] <String[]>] [-NoHeader] [-DataOnly] [-ConvertEmptyStringsToNull] [-UseMSSQLSyntax] [<CommonParameters>]
```

## DESCRIPTION

Generate SQL insert statements from Excel spreadsheet.

## EXAMPLES

### EXAMPLE 1

```text
Generate SQL insert statements from Movies.xlsx file, leaving blank cells as empty strings:

----------------------------------------------------------
| File: Movies.xlsx            -           Sheet: Sheet1 |
----------------------------------------------------------
|           A            B            C                  |
|1     Movie Name      Year         Rating               |
|2     The Bodyguard   1992           9                  |
|3     The Matrix      1999           8                  |
|4     Skyfall         2012           9                  |
|5     The Avengers    2012                              |
----------------------------------------------------------

PS C:\>  ConvertFrom-ExcelToSQLInsert -TableName "Movies" -Path 'C:\Movies.xlsx'
INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('The Bodyguard', '1992', '9');
INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('The Matrix', '1999', '8');
INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('Skyfall', '2012', '9');
INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('The Avengers', '2012', '');
```

### EXAMPLE 2

```text
Generate SQL insert statements from Movies.xlsx file, specify NULL instead of an empty string.

----------------------------------------------------------
| File: Movies.xlsx            -           Sheet: Sheet1 |
----------------------------------------------------------
|           A            B            C                  |
|1     Movie Name      Year         Rating               |
|2     The Bodyguard   1992           9                  |
|3     The Matrix      1999           8                  |
|4     Skyfall         2012           9                  |
|5     The Avengers    2012                              |
----------------------------------------------------------

PS C:\> ConvertFrom-ExcelToSQLInsert -TableName "Movies" -Path "C:\Movies.xlsx" -ConvertEmptyStringsToNull
INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('The Bodyguard', '1992', '9');
INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('The Matrix', '1999', '8');
INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('Skyfall', '2012', '9');
INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('The Avengers', '2012', NULL);
```

## PARAMETERS

### -TableName

Name of the target database table.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path

Path to an existing .XLSX file This parameter is passed to Import-Excel as is.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: FullName

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -WorkSheetname

Specifies the name of the worksheet in the Excel workbook to import. By default, if no name is provided, the first worksheet will be imported. This parameter is passed to Import-Excel as is.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Sheet

Required: False
Position: 3
Default value: 1
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartRow

The row from where we start to import data, all rows above the StartRow are disregarded. By default this is the first row. When the parameters '-NoHeader' and '-HeaderName' are not provided, this row will contain the column headers that will be used as property names. When one of both parameters are provided, the property names are automatically created and this row will be treated as a regular row containing data.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: HeaderRow, TopRow

Required: False
Position: 4
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header

Specifies custom property names to use, instead of the values defined in the column headers of the TopRow. If you provide fewer header names than there is data in the worksheet, then only the data with a corresponding header name will be imported and the data without header name will be disregarded. If you provide more header names than there is data in the worksheet, then all data will be imported and all objects will have all the property names you defined in the header names. As such, the last properties will be blank as there is no data for them.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader

Automatically generate property names \(P1, P2, P3, ..\) instead of the ones defined in the column headers of the TopRow. This switch is best used when you want to import the complete worksheet 'as is' and are not concerned with the property names.

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

### -DataOnly

Import only rows and columns that contain data, empty rows and empty columns are not imported.

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

### -ConvertEmptyStringsToNull

If specified, cells without any data are replaced with NULL, instead of an empty string. This is to address behviors in certain DBMS where an empty string is insert as 0 for INT column, instead of a NULL value.

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

### -UseMSSQLSyntax

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

