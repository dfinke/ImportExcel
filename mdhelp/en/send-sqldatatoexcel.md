---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Send-SQLDataToExcel

## SYNOPSIS

Inserts a DataTable - returned by a SQL query - into an ExcelSheet

## SYNTAX

### none \(Default\)

```text
Send-SQLDataToExcel [-QueryTimeout <Int32>] [-Force] [-Path <String>] [-ExcelPackage <ExcelPackage>]
 [-Calculate] [-Show] [-WorksheetName <String>] [-Password <String>] [-ClearSheet] [-Append] [-Title <String>]
 [-TitleFillPattern <ExcelFillStyle>] [-TitleBold] [-TitleSize <Int32>] [-TitleBackgroundColor <Object>]
 [-IncludePivotTable] [-PivotTableName <String>] [-PivotRows <String[]>] [-PivotColumns <String[]>]
 [-PivotData <Object>] [-PivotFilter <String[]>] [-PivotDataToColumn] [-PivotTableDefinition <Hashtable>]
 [-IncludePivotChart] [-ChartType <eChartType>] [-NoLegend] [-ShowCategory] [-ShowPercent] [-AutoSize]
 [-MaxAutoSizeRows <Object>] [-NoClobber] [-FreezeTopRow] [-FreezeFirstColumn] [-FreezeTopRowFirstColumn]
 [-FreezePane <Int32[]>] [-AutoFilter] [-BoldTopRow] [-NoHeader] [-RangeName <String>] [-TableName <Object>]
 [-TableStyle <TableStyles>] [-Barchart] [-PieChart] [-LineChart] [-ColumnChart]
 [-ExcelChartDefinition <Object[]>] [-HideSheet <String[]>] [-UnHideSheet <String[]>] [-MoveToStart]
 [-MoveToEnd] [-MoveBefore <Object>] [-MoveAfter <Object>] [-KillExcel] [-AutoNameRange] [-StartRow <Int32>]
 [-StartColumn <Int32>] [-PassThru] [-Numberformat <String>] [-ExcludeProperty <String[]>]
 [-NoAliasOrScriptPropeties] [-DisplayPropertySet] [-NoNumberConversion <String[]>]
 [-ConditionalFormat <Object[]>] [-ConditionalText <Object[]>] [-Style <Object[]>] [-CellStyleSB <ScriptBlock>]
 [-Activate] [-Now] [-ReturnRange] [-PivotTotals <String>] [-NoTotalsInPivot] [-ReZip] [<CommonParameters>]
```

### ODBCConnection

```text
Send-SQLDataToExcel -Connection <Object> -SQL <String> [-QueryTimeout <Int32>] [-Force] [-Path <String>]
 [-ExcelPackage <ExcelPackage>] [-Calculate] [-Show] [-WorksheetName <String>] [-Password <String>]
 [-ClearSheet] [-Append] [-Title <String>] [-TitleFillPattern <ExcelFillStyle>] [-TitleBold]
 [-TitleSize <Int32>] [-TitleBackgroundColor <Object>] [-IncludePivotTable] [-PivotTableName <String>]
 [-PivotRows <String[]>] [-PivotColumns <String[]>] [-PivotData <Object>] [-PivotFilter <String[]>]
 [-PivotDataToColumn] [-PivotTableDefinition <Hashtable>] [-IncludePivotChart] [-ChartType <eChartType>]
 [-NoLegend] [-ShowCategory] [-ShowPercent] [-AutoSize] [-MaxAutoSizeRows <Object>] [-NoClobber]
 [-FreezeTopRow] [-FreezeFirstColumn] [-FreezeTopRowFirstColumn] [-FreezePane <Int32[]>] [-AutoFilter]
 [-BoldTopRow] [-NoHeader] [-RangeName <String>] [-TableName <Object>] [-TableStyle <TableStyles>] [-Barchart]
 [-PieChart] [-LineChart] [-ColumnChart] [-ExcelChartDefinition <Object[]>] [-HideSheet <String[]>]
 [-UnHideSheet <String[]>] [-MoveToStart] [-MoveToEnd] [-MoveBefore <Object>] [-MoveAfter <Object>]
 [-KillExcel] [-AutoNameRange] [-StartRow <Int32>] [-StartColumn <Int32>] [-PassThru] [-Numberformat <String>]
 [-ExcludeProperty <String[]>] [-NoAliasOrScriptPropeties] [-DisplayPropertySet]
 [-NoNumberConversion <String[]>] [-ConditionalFormat <Object[]>] [-ConditionalText <Object[]>]
 [-Style <Object[]>] [-CellStyleSB <ScriptBlock>] [-Activate] [-Now] [-ReturnRange] [-PivotTotals <String>]
 [-NoTotalsInPivot] [-ReZip] [<CommonParameters>]
```

### SQLConnection

```text
Send-SQLDataToExcel -Connection <Object> [-MsSQLserver] [-DataBase <String>] -SQL <String> [-QueryTimeout <Int32>] [-Force] [-Path <String>] [-ExcelPackage <ExcelPackage>] [-Calculate] [-Show] [-WorksheetName <String>] [-Password <String>] [-ClearSheet] [-Append] [-Title <String>] [-TitleFillPattern <ExcelFillStyle>] [-TitleBold] [-TitleSize <Int32>] [-TitleBackgroundColor <Object>] [-IncludePivotTable] [-PivotTableName <String>] [-PivotRows <String[]>] [-PivotColumns <String[]>] [-PivotData <Object>] [-PivotFilter <String[]>] [-PivotDataToColumn] [-PivotTableDefinition <Hashtable>] [-IncludePivotChart] [-ChartType <eChartType>] [-NoLegend] [-ShowCategory] [-ShowPercent] [-AutoSize] [-MaxAutoSizeRows <Object>] [-NoClobber] [-FreezeTopRow] [-FreezeFirstColumn] [-FreezeTopRowFirstColumn] [-FreezePane <Int32[]>] [-AutoFilter] [-BoldTopRow] [-NoHeader] [-RangeName <String>] [-TableName <Object>] [-TableStyle <TableStyles>] [-Barchart] [-PieChart] [-LineChart] [-ColumnChart] [-ExcelChartDefinition <Object[]>] [-HideSheet <String[]>] [-UnHideSheet <String[]>] [-MoveToStart] [-MoveToEnd] [-MoveBefore <Object>] [-MoveAfter <Object>] [-KillExcel] [-AutoNameRange] [-StartRow <Int32>] [-StartColumn <Int32>] [-PassThru] [-Numberformat <String>] [-ExcludeProperty <String[]>] [-NoAliasOrScriptPropeties] [-DisplayPropertySet] [-NoNumberConversion <String[]>] [-ConditionalFormat <Object[]>] [-ConditionalText <Object[]>] [-Style <Object[]>] [-CellStyleSB <ScriptBlock>] [-Activate] [-Now] [-ReturnRange] [-PivotTotals <String>] [-NoTotalsInPivot] [-ReZip] [<CommonParameters>]
```

### ExistingSession

```text
Send-SQLDataToExcel -Session <Object> -SQL <String> [-QueryTimeout <Int32>] [-Force] [-Path <String>] [-ExcelPackage <ExcelPackage>] [-Calculate] [-Show] [-WorksheetName <String>] [-Password <String>] [-ClearSheet] [-Append] [-Title <String>] [-TitleFillPattern <ExcelFillStyle>] [-TitleBold] [-TitleSize <Int32>] [-TitleBackgroundColor <Object>] [-IncludePivotTable] [-PivotTableName <String>] [-PivotRows <String[]>] [-PivotColumns <String[]>] [-PivotData <Object>] [-PivotFilter <String[]>] [-PivotDataToColumn] [-PivotTableDefinition <Hashtable>] [-IncludePivotChart] [-ChartType <eChartType>] [-NoLegend] [-ShowCategory] [-ShowPercent] [-AutoSize] [-MaxAutoSizeRows <Object>] [-NoClobber] [-FreezeTopRow] [-FreezeFirstColumn] [-FreezeTopRowFirstColumn] [-FreezePane <Int32[]>] [-AutoFilter] [-BoldTopRow] [-NoHeader] [-RangeName <String>] [-TableName <Object>] [-TableStyle <TableStyles>] [-Barchart] [-PieChart] [-LineChart] [-ColumnChart] [-ExcelChartDefinition <Object[]>] [-HideSheet <String[]>] [-UnHideSheet <String[]>] [-MoveToStart] [-MoveToEnd] [-MoveBefore <Object>] [-MoveAfter <Object>] [-KillExcel] [-AutoNameRange] [-StartRow <Int32>] [-StartColumn <Int32>] [-PassThru] [-Numberformat <String>] [-ExcludeProperty <String[]>] [-NoAliasOrScriptPropeties] [-DisplayPropertySet] [-NoNumberConversion <String[]>] [-ConditionalFormat <Object[]>] [-ConditionalText <Object[]>] [-Style <Object[]>] [-CellStyleSB <ScriptBlock>] [-Activate] [-Now] [-ReturnRange] [-PivotTotals <String>] [-NoTotalsInPivot] [-ReZip] [<CommonParameters>]
```

### Pre-FetchedData

```text
Send-SQLDataToExcel [-QueryTimeout <Int32>] -DataTable <DataTable> [-Force] [-Path <String>] [-ExcelPackage <ExcelPackage>] [-Calculate] [-Show] [-WorksheetName <String>] [-Password <String>] [-ClearSheet] [-Append] [-Title <String>] [-TitleFillPattern <ExcelFillStyle>] [-TitleBold] [-TitleSize <Int32>] [-TitleBackgroundColor <Object>] [-IncludePivotTable] [-PivotTableName <String>] [-PivotRows <String[]>] [-PivotColumns <String[]>] [-PivotData <Object>] [-PivotFilter <String[]>] [-PivotDataToColumn] [-PivotTableDefinition <Hashtable>] [-IncludePivotChart] [-ChartType <eChartType>] [-NoLegend] [-ShowCategory] [-ShowPercent] [-AutoSize] [-MaxAutoSizeRows <Object>] [-NoClobber] [-FreezeTopRow] [-FreezeFirstColumn] [-FreezeTopRowFirstColumn] [-FreezePane <Int32[]>] [-AutoFilter] [-BoldTopRow] [-NoHeader] [-RangeName <String>] [-TableName <Object>] [-TableStyle <TableStyles>] [-Barchart] [-PieChart] [-LineChart] [-ColumnChart] [-ExcelChartDefinition <Object[]>] [-HideSheet <String[]>] [-UnHideSheet <String[]>] [-MoveToStart] [-MoveToEnd] [-MoveBefore <Object>] [-MoveAfter <Object>] [-KillExcel] [-AutoNameRange] [-StartRow <Int32>] [-StartColumn <Int32>] [-PassThru] [-Numberformat <String>] [-ExcludeProperty <String[]>] [-NoAliasOrScriptPropeties] [-DisplayPropertySet] [-NoNumberConversion <String[]>] [-ConditionalFormat <Object[]>] [-ConditionalText <Object[]>] [-Style <Object[]>] [-CellStyleSB <ScriptBlock>] [-Activate] [-Now] [-ReturnRange] [-PivotTotals <String>] [-NoTotalsInPivot] [-ReZip] [<CommonParameters>]
```

## DESCRIPTION

This command takes a SQL statement and run it against a database connection; for the connection it accepts either

* an object representing a session with a SQL server or ODBC database, or
* a connection string to make a session \(if -MSSQLServer is specified it uses the SQL Native client,

  and -Connection can be a server name instead of a detailed connection string. Without this switch it uses ODBC\)

The command takes all the parameters of Export-Excel, except for -InputObject \(alias TargetData\); after fetching the data it calls Export-Excel with the data as the value of InputParameter and whichever of Export-Excel's parameters it was passed; for details of these parameters see the help for Export-Excel.

## EXAMPLES

### EXAMPLE 1

```text
PS\> Send-SQLDataToExcel -MsSQLserver -Connection localhost -SQL  "select name,type,type_desc from [master].[sys].[all_objects]" -Path .\temp.xlsx -WorkSheetname master -AutoSize -FreezeTopRow -AutoFilter -BoldTopRow
```

Connects to the local SQL server and selects 3 columns from \[Sys\].\[all\_objects\] and exports then to a sheet named master with some basic header management

### EXAMPLE 2

```text
PS\> $dbPath = 'C:\Users\James\Documents\Database1.accdb'
PS\> $Connection = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=$dbPath;"
PS\> $SQL="SELECT top 25 Name,Length  From TestData ORDER BY Length DESC"
PS\> Send-SQLDataToExcel -Connection  $connection -SQL $sql -path .\demo1.xlsx -WorkSheetname "Sizes" -AutoSize
```

This creates an ODBC connection string to read from an Access file and a SQL Statement to extracts data from it, and sends the resulting data to a new worksheet

### EXAMPLE 3

```text
PS\> $dbPath = 'C:\users\James\Documents\f1Results.xlsx'

PS\> $Connection = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=$dbPath;"
PS\> $SQL="SELECT top 25 DriverName, Count(RaceDate) as Races, Count(Win) as Wins, Count(Pole) as Poles, Count(FastestLap) as Fastlaps " +
           " FROM Results  GROUP BY DriverName ORDER BY (count(win)) DESC"

PS\>Send-SQLDataToExcel -Connection  $connection -SQL $sql -path .\demo2.xlsx -WorkSheetname "Winners" -AutoSize -AutoNameRange -ConditionalFormat @{DataBarColor="Blue"; Range="Wins"}
```

Similar to the previous example, this creates a connection string, this time for an Excel file, and runs a SQL statement to get a list of motor-racing results, outputting the resulting data to a new spreadsheet. The spreadsheet is formatted and a data bar added to show make the drivers' wins clearer. \(The F1 results database is available from [https://1drv.ms/x/s!AhfYu7-CJv4ehNdZWxJE9LMAX\_N5sg](https://1drv.ms/x/s!AhfYu7-CJv4ehNdZWxJE9LMAX_N5sg) \)

### EXAMPLE 4

```text
PS\> $dbPath = 'C:\users\James\Documents\f1Results.xlsx'

PS\> $SQL = "SELECT top 25 DriverName, Count(RaceDate) as Races, Count(Win) as Wins, Count(Pole) as Poles, Count(FastestLap) as Fastlaps " +
             " FROM Results GROUP BY DriverName ORDER BY (count(win)) DESC"
PS\> $null = Get-SQL -Session F1 -excel -Connection $dbPath -sql $sql -OutputVariable Table

PS\> Send-SQLDataToExcel -DataTable $Table -Path ".\demo3.xlsx" -WorkSheetname Gpwinners -autosize  -TableName winners -TableStyle Light6 -show
```

This uses Get-SQL \(at least V1.1 - download from the PowerShell gallery with Install-Module -Name GetSQL - \(note the function is Get-SQL the module is GetSQL without the "-" \)

Get-SQL simplifies making database connections and building /submitting SQL statements. Here Get-SQL uses the same SQL statement as before; -OutputVariable leaves a System.Data.DataTable object in $table and Send-SQLDataToExcel puts $table into the worksheet and sets it as an Excel table. The command is equivalent to running

PS&gt; Export-Excel -inputObject $Table -Path ".\demo3.xlsx" -WorkSheetname Gpwinners -autosize -TableName winners -TableStyle Light6 -show

This is quicker than using PS&gt; Get-SQL \ \| export-excel -ExcludeProperty rowerror,rowstate,table,itemarray,haserrors \

\(the F1 results database is available from [https://1drv.ms/x/s!AhfYu7-CJv4ehNdZWxJE9LMAX\_N5sg](https://1drv.ms/x/s!AhfYu7-CJv4ehNdZWxJE9LMAX_N5sg) \)

### EXAMPLE 5

```text
PS\>$SQL = "SELECT top 25 DriverName,  Count(Win) as Wins FROM Results GROUP BY DriverName ORDER BY (count(win)) DESC"
PS\> Send-SQLDataToExcel -Session $DbSessions\["f1"\] -SQL $sql -Path  ".\demo3.xlsx" -WorkSheetname Gpwinners  -ClearSheet -autosize -ColumnChart
```

Like the previous example, this uses Get-SQL \(download from the gallery with Install-Module -Name GetSQL\). It uses the database session which Get-SQL created, rather than an ODBC connection string. The Session parameter can either be a object \(as shown here\), or the name used by Get-SQL \("F1" in this case\).

Here the data is presented as a quick chart.

### EXAMPLE 6

```text
Send-SQLDataToExcel -path .\demo4.xlsx -WorkSheetname "LR" -Connection "DSN=LR" -sql "SELECT name AS CollectionName FROM AgLibraryCollection Collection ORDER BY CollectionName"
```

This example uses an Existing ODBC datasource name "LR" which maps to an adobe lightroom database and gets a list of collection names into a worksheet

## PARAMETERS

### -Connection

A database connection string to be used to create a database session; either

* A Data source name written in the form DSN=ODBC\_Data\_Source\_Name, or
* A full ODBC or SQL Native Client Connection string, or
* The name of a SQL server.

```yaml
Type: Object
Parameter Sets: ODBCConnection, SQLConnection
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Session

An active ODBC Connection or SQL connection object representing a session with a database which will be queried to get the data .

```yaml
Type: Object
Parameter Sets: ExistingSession
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MsSQLserver

Specifies the connection string is for SQL server, not ODBC.

```yaml
Type: SwitchParameter
Parameter Sets: SQLConnection
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -DataBase

Switches to a specific database on a SQL server.

```yaml
Type: String
Parameter Sets: SQLConnection
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SQL

The SQL query to run against the session which was passed in -Session or set up from -Connection.

```yaml
Type: String
Parameter Sets: ODBCConnection, SQLConnection, ExistingSession
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -QueryTimeout

Override the default query time of 30 seconds.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -DataTable

A System.Data.DataTable object containing the data to be inserted into the spreadsheet without running a query. This remains supported to avoid breaking older scripts, but if you have a DataTable object you can pass the it into Export-Excel using -InputObject.

```yaml
Type: DataTable
Parameter Sets: Pre-FetchedData
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force

If specified Export-Excel will be called with parameters specified, even if there is no data to send

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

[Export-Excel](send-sqldatatoexcel.md)

