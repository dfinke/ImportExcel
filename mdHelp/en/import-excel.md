---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: 'https://github.com/dfinke/ImportExcel'
schema: 2.0.0
---

# Import-Excel

## SYNOPSIS

Create custom objects from the rows in an Excel worksheet.

## SYNTAX

### PathC

```text
Import-Excel [-Path] <String> [[-WorksheetName] <String>] [-NoHeader] [-StartRow <Int32>] [-EndRow <Int32>] [-StartColumn <Int32>] [-EndColumn <Int32>] [-DataOnly] [-AsText <String[]>] [-Password <String>]  [<CommonParameters>]
```

### PathB

```text
Import-Excel [-Path] <String> [[-WorksheetName] <String>] -HeaderName <String[]> [-StartRow <Int32>] [-EndRow <Int32>] [-StartColumn <Int32>] [-EndColumn <Int32>] [-DataOnly] [-AsText <String[]>] [-Password <String>] [<CommonParameters>]
```

### PathA

```text
Import-Excel [-Path] <String> [[-WorksheetName] <String>] [-StartRow <Int32>] [-EndRow <Int32>] [-StartColumn <Int32>] [-EndColumn <Int32>] [-DataOnly] [-AsText <String[]>] [-Password <String>] [<CommonParameters>]
```

### PackageC

```text
Import-Excel -ExcelPackage <ExcelPackage> [[-WorksheetName] <String>] [-NoHeader] [-StartRow <Int32>] [-EndRow <Int32>] [-StartColumn <Int32>] [-EndColumn <Int32>] [-DataOnly] [-AsText <String[]>] [-Password <String>] [<CommonParameters>]
```

### PackageB

```text
Import-Excel -ExcelPackage <ExcelPackage> [[-WorksheetName] <String>] -HeaderName <String[]> [-StartRow <Int32>] [-EndRow <Int32>] [-StartColumn <Int32>] [-EndColumn <Int32>] [-DataOnly] [-AsText <String[]>] [-Password <String>] [<CommonParameters>]
```

### PackageA

```text
Import-Excel -ExcelPackage <ExcelPackage> [[-WorksheetName] <String>] [-StartRow <Int32>] [-EndRow <Int32>] [-StartColumn <Int32>] [-EndColumn <Int32>] [-DataOnly] [-AsText <String[]>] [-Password <String>] [<CommonParameters>]
```

## DESCRIPTION

The Import-Excel cmdlet creates custom objects from the rows in an Excel worksheet. Each row is represented as one object.

This is possible without installing Microsoft Excel by using the .NET library 'EPPLus.dll'.

By default, the property names of the objects are retrieved from the column headers. Because an object cannot have a blank property name, only columns with column headers will be imported.

If the default behavior is not desired and you want to import the complete worksheet 'as is', the parameter '-NoHeader' can be used. In case you want to provide your own property names, you can use the parameter '-HeaderName'.

## EXAMPLES

### EXAMPLE 1

```text

----------------------------------------------
| File: Movies.xlsx     -      Sheet: Actors |
----------------------------------------------
|           A           B            C       |
|1     First Name                 Address    |
|2     Chuck         Norris       California |
|3     Jean-Claude   Vandamme     Brussels   |
----------------------------------------------

PS C:> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Actors

First Name: Chuck
Address   : California

First Name: Jean-Claude
Address   : Brussels
```

Import data from an Excel worksheet. One object is created for each row. The property names of the objects consist of the column names defined in the first row. In case a column doesn't have a column header \(usually in row 1 when '-StartRow' is not used\), then the unnamed columns will be skipped and the data in those columns will not be imported.

Notice that column 'B' is not imported because there's no value in cell 'B1' that can be used as property name for the objects.

### EXAMPLE 2

```text

----------------------------------------------
| File: Movies.xlsx     -      Sheet: Actors |
----------------------------------------------
|           A           B            C       |
|1     First Name                 Address    |
|2     Chuck         Norris       California |
|3     Jean-Claude   Vandamme     Brussels   |
----------------------------------------------

PS\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Actors -NoHeader

P1: First Name
P2:
P3: Address

P1: Chuck
P2: Norris
P3: California

P1: Jean-Claude
P2: Vandamme
P3: Brussels
```

Imports the complete Excel worksheet 'as is' by using the '-NoHeader' switch. One object is created for each row. The property names of the objects will be automatically generated \(P1, P2, P3, ..\).

Notice that the column header \(row 1\) is imported as an object too.

### EXAMPLE 3

```text

----------------------------------------------------------
| File: Movies.xlsx            -           Sheet: Movies |
----------------------------------------------------------
|           A            B            C          D       |
|1     The Bodyguard   1992           9                  |
|2     The Matrix      1999           8                  |
|3                                                       |
|4     Skyfall         2012           9                  |
----------------------------------------------------------

PS\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Movies -HeaderName 'Movie name', 'Year', 'Rating', 'Genre'

Movie name: The Bodyguard
Year      : 1992
Rating    : 9
Genre     :

Movie name: The Matrix
Year      : 1999
Rating    : 8
Genre     :

Movie name:
Year      :
Rating    :
Genre     :

Movie name: Skyfall
Year      : 2012
Rating    : 9
Genre     :
```

This imports data from an Excel worksheet, and as before one object is created for each row. The property names for the objects are defined in the parameter '-HeaderName'. The properties are named starting from the most left column \(A\) to the right. In case no value is present in one of the columns, that property will have an empty value.

Notice that empty rows are imported and that data for the property 'Genre' is not present in the worksheet. As such, the 'Genre' property will be blank for all objects.

### EXAMPLE 4

```text

----------------------------------------------------------
| File: Movies.xlsx            -           Sheet: Movies |
----------------------------------------------------------
|           A            B            C          D       |
|1     The Bodyguard   1992           9                  |
|2     The Matrix      1999           8                  |
|3                                                       |
|4     Skyfall         2012           9                  |
----------------------------------------------------------

PS\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Movies -NoHeader -DataOnly

P1: The Bodyguard
P2: 1992
P3: 9

P1: The Matrix
P2: 1999
P3: 8

P1: Skyfall
P2: 2012
P3: 9
```

Import data from an Excel worksheet, and one object is created for each non-blank row. The property names of the objects \(P1, P2, P3, ..\) are automatically generated by using the switch '-NoHeader' . The switch '-DataOnly' will speed up the import because empty rows and empty columns are not imported.

Notice that empty rows and empty columns are not imported.

### EXAMPLE 5

```text

----------------------------------------------------------
| File: Movies.xlsx            -           Sheet: Actors |
----------------------------------------------------------
|           A           B           C            D       |
|1     Chuck                     Norris       California |
|2                                                       |
|3     Jean-Claude               Vandamme     Brussels   |
----------------------------------------------------------

PS\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Actors -DataOnly -HeaderName 'FirstName', 'SecondName', 'City' -StartRow 2

FirstName : Jean-Claude
SecondName: Vandamme
City      : Brussels
```

Import data from an Excel worksheet. One object is created for each row. The property names are provided with the '-HeaderName' parameter. The import will start from row 2 and empty columns and rows are not imported.

Notice that only 1 object is imported with only 3 properties. Column B and row 2 are empty and have been disregarded by using the switch '-DataOnly'. The property names have been named with the values provided with the parameter '-HeaderName'. Row number 1 with 'Chuck Norris' has not been imported, because we started the import from row 2 with the parameter '-StartRow 2'.

### EXAMPLE 6

```text
PS\> ,(Import-Excel -Path .\SysTables_AdventureWorks2014.xlsx) |
    Write-SqlTableData -ServerInstance localhost\DEFAULT -Database BlankDB -SchemaName dbo -TableName MyNewTable_fromExcel -Force
```

Imports data from an Excel file and pipes the result to the Write-SqlTableData command to be INSERTed into a table in a SQL Server database.

The ",\( ... \)" around the Import-Excel command allows all rows to be imported from the Excel file, prior to pipelining to the Write-SqlTableData cmdlet. This helps prevent a RBAR scenario and is important when trying to import thousands of rows.

The -Force parameter will be ignored if the table already exists. However, if a table is not found that matches the values provided by -SchemaName and -TableName parameters, it will create a new table in SQL Server database.

The Write-SqlTableData cmdlet will inherit the column names & datatypes for the new table from the object being piped in. NOTE: You need to install the SqlServer module from the PowerShell Gallery in oder to get the Write-SqlTableData cmdlet.

## PARAMETERS

### -Path

Specifies the path to the Excel file.

```yaml
Type: String
Parameter Sets: PathC, PathB, PathA
Aliases: FullName

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -ExcelPackage

Instead of specifying a path, provides an Excel Package object \(from Open-ExcelPackage\). Using this avoids re-reading the whole file when importing multiple parts of it.

To allow multiple read operations Import-Excel does NOT close the package, and you should use Close-ExcelPackage -noSave to close it.

```yaml
Type: ExcelPackage
Parameter Sets: PackageC, PackageB, PackageA
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WorksheetName

Specifies the name of the worksheet in the Excel workbook to import. By default, if no name is provided, the first worksheet will be imported.

```yaml
Type: String
Parameter Sets: (All)
Aliases: Sheet

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderName

Specifies custom property names to use, instead of the values defined in the column headers of the TopRow.

If you provide fewer header names than there are columns of data in the worksheet, then data will only be imported from that number of columns - the others will be ignored.

If you provide more header names than there are columns of data in the worksheet, it will result in blank properties being added to the objects returned.

```yaml
Type: String[]
Parameter Sets: PathB, PackageB
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader

Automatically generate property names \(P1, P2, P3, ..\) instead of the ones defined in the column headers of the TopRow.

This switch is best used when you want to import the complete worksheet 'as is' and are not concerned with the property names.

```yaml
Type: SwitchParameter
Parameter Sets: PathC, PackageC
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartRow

The row from where we start to import data, all rows above the StartRow are disregarded. By default this is the first row.

When the parameters '-NoHeader' and '-HeaderName' are not provided, this row will contain the column headers that will be used as property names. If either is provided, the property names are automatically created and this row will be treated as a regular row containing data.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: HeaderRow, TopRow

Required: False
Position: Named
Default value: 1
Accept pipeline input: False
Accept wildcard characters: False
```

### -EndRow

By default all rows up to the last cell in the sheet will be imported. If specified, import stops at this row.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: StopRow, BottomRow

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartColumn

The number of the first column to read data from \(1 by default\).

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: LeftColumn

Required: False
Position: Named
Default value: 1
Accept pipeline input: False
Accept wildcard characters: False
```

### -EndColumn

By default the import reads up to the last populated column, -EndColumn tells the import to stop at an earlier number.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: RightColumn

Required: False
Position: Named
Default value: 0
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

### -AsText

Normally Import-Excel returns the Cell values. AsText allows selected columns to be returned as the text displayed in their cells. \(\* is supported as a wildcard.\)

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
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
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password

Accepts a string that will be used to open a password protected Excel file.

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

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about\_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS

[https://github.com/dfinke/ImportExcel](https://github.com/dfinke/ImportExcel)

