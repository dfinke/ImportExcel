---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: https://dfinke.github.io/ImportExcel/commands/Read-OleDbData
schema: 2.0.0
---

# Read-OleDbData

## SYNOPSIS
Read data from an OleDb source using dotnet classes.
This allows for OleDb queries against excel spreadsheets.
Examples will only be for querying xlsx files.

For additional documentation, see Microsoft's documentation on the System.Data OleDb namespace here:
https://docs.microsoft.com/en-us/dotnet/api/system.data.oledb

## SYNTAX

```
Read-OleDbData [-ConnectionString] <String> [-SqlStatement] <String> [<CommonParameters>]
```

## DESCRIPTION
Read data from an OleDb source using dotnet classes.
This allows for OleDb queries against excel spreadsheets.
Examples will only be for querying xlsx files using ACE.

## EXAMPLES

### EXAMPLE 1
```
Read-OleDbData `
```

-ConnectionString "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=file.xlsx;Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'" \`
    -SqlStatement "select ROUND(F1) as \[A\] from \[sheet1$A1:A1\]"

### EXAMPLE 2
```
$ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=file.xlsx;Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'"
```

$SqlStatement = "select ROUND(F1) as \[A\] from \[sheet1$A1:A1\]"
Read-OleDbData -ConnectionString $ConnectionString -SqlStatement $SqlStatement

### EXAMPLE 3
```
$ReadDataArgs = @{
```

SqlStatement = Get-Content query.sql -Raw
    ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=file.xlsx;Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'"
}
$Results = Read-OleDbData @ReadDataArgs

## PARAMETERS

### -ConnectionString
{{ Fill ConnectionString Description }}

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

### -SqlStatement
{{ Fill SqlStatement Description }}

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
