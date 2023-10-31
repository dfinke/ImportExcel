---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: https://dfinke.github.io/ImportExcel/commands/Invoke-ExcelQuery
schema: 2.0.0
---

# Invoke-ExcelQuery

## SYNOPSIS
Helper method for executing Read-OleDbData with some basic defaults.

For additional help, see documentation for Read-OleDbData cmdlet.

## SYNTAX

```
Invoke-ExcelQuery [-Path] <String> [-Query] <String> [<CommonParameters>]
```

## DESCRIPTION
Uses Read-OleDbData to execute a sql statement against a xlsx file.
For finer grained control over the interaction, you may use that cmdlet.
This cmdlet assumes a file path will be passed in and the connection string will be built with no headers and treating all results as text.

Running this command is equivalent to running the following:

$FullName = (Get-ChildItem $Path).FullName
Read-OleDbData \`
    -ConnectionString "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$FullName;Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'" \`
    -SqlStatement $Query

Note that this command uses the MICROSOFT.ACE.OLEDB provider and will not work without it.

If needed, please download the appropriate package from https://www.microsoft.com/en-us/download/details.aspx?id=54920.

## EXAMPLES

### EXAMPLE 1
```
Invoke-ExcelQuery .\test.xlsx 'select ROUND(F1) as [A1] from [sheet3$A1:A1]'
```

### EXAMPLE 2
```
$Path = (Get-ChildItem 'test.xlsx').FullName
```

$Query = "select ROUND(F1) as \[A\] from \[sheet1$A1:A1\]"
Read-XlsxUsingOleDb -Path $Path -Query $Query

### EXAMPLE 3
```
$ReadDataArgs = @{
```

Path = .\test.xlsx
    Query = Get-Content query.sql -Raw
}
$Results = Invoke-ExcelQuery @ReadDataArgs

## PARAMETERS

### -Path
The path to the file to open.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Query
var name consistent with Import-Excel
var name consistent with Invoke-Sqlcmd

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
