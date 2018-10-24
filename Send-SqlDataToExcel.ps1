Function Send-SQLDataToExcel {
    <#
        .SYNOPSIS
            Inserts a DataTable - returned by SQL query into an ExcelSheet, more efficiently than sending it via Export-Excel
        .DESCRIPTION
            This command can accept a data table object or take a SQL statement and run it against a database connection.
            If running a SQL statement, the accepts either
            * an object representing a session with a SQL server or ODBC database, or
            * a connection String to make a session.
            The command takes most of the parameters of Export-Excel, and after inserting the table into the worksheet it
            calls Export-Excel to carry out other tasks on the sheet. It is more efficient to do this than to get data-rows
             and pipe them into Export-Excel, stripped off the database 'housekeeping' properties.
        .PARAMETER DataTable
            A System.Data.DataTable object containing the data to be inserted into the spreadsheet without running a query.
        .PARAMETER Session
            An active ODBC Connection or SQL connection object representing a session with a database which will be queried to get the data .
        .PARAMETER Connection
            A database connection string to be used to create a database session; either
            * A Data source name written in the form DSN=ODBC_Data_Source_Name, or
            * A full odbc or SQL Connection string, or
            * The name of a SQL server.
        .PARAMETER MSSQLServer
            Specifies the connection string is for SQL server, not ODBC.
        .PARAMETER SQL
            The SQL query to run against the session which was passed in -Session or set up from -Connection.
        .PARAMETER Database
           Switches to a specific database on a SQL server.
        .PARAMETER QueryTimeout
            Override the default query time of 30 seconds.
        .PARAMETER Path
            Path to a new or existing .XLSX file.
        .PARAMETER WorkSheetName
            The name of a sheet within the workbook - "Sheet1" by default.
        .PARAMETER KillExcel
            Closes Excel - prevents errors writing to the file because Excel has it open.
        .PARAMETER Title
            Text of a title to be placed in the top left cell.
        .PARAMETER TitleBold
            Sets the title in boldface type.
        .PARAMETER TitleSize
            Sets the point size for the title.
        .PARAMETER TitleBackgroundColor
            Sets the cell background color for the title cell.
        .PARAMETER TitleFillPattern
            Sets the fill pattern for the title cell.
        .PARAMETER Password
            Sets password protection on the workbook.
        .PARAMETER IncludePivotTable
            Adds a Pivot table using the data in the worksheet.
        .PARAMETER PivotTableName
            If a Pivot table is created from command line parameters, specificies the name of the new sheet holding the pivot. If Omitted this will be "WorksheetName-PivotTable"
        .PARAMETER PivotRows
            Name(s) columns from the spreadhseet which will provide the Row name(s) in a pivot table created from command line parameters.
        .PARAMETER PivotColumns
            Name(s) columns from the spreadhseet which will provide the Column name(s) in a pivot table created from command line parameters.
        .PARAMETER PivotFilter
            Name(s) columns from the spreadhseet which will provide the Filter name(s) in a pivot table created from command line parameters.
        .PARAMETER PivotData
            In a pivot table created from command line parameters, the fields to use in the table body is given as a Hash table in the form ColumnName = Average|Count|CountNums|Max|Min|Product|None|StdDev|StdDevP|Sum|Var|VarP .
        .PARAMETER PivotDataToColumn
            If there are multiple datasets in a PivotTable, by default they are shown seperatate rows under the given row heading; this switch makes them seperate columns.
        .PARAMETER NoTotalsInPivot
            In a pivot table created from command line parameters, prevents the addition of totals to rows and columns.
        .PARAMETER IncludePivotChart
            Include a chart with the Pivot table - implies -IncludePivotTable.
        .PARAMETER ChartType
            The type for Pivot chart (one of Excel's defined chart types)
        .PARAMETER NoLegend
            Exclude the legend from the pivot chart.
        .PARAMETER ShowCategory
            Add category labels to the pivot chart.
        .PARAMETER ShowPercent
            Add Percentage labels to the pivot chart.
        .PARAMETER PivotTableDefinition
            Instead of describing a single pivot table with mutliple commandline paramters; you can use a HashTable in the form PivotTableName = Definition;
            Definition is itself a hashtable with Sheet PivotTows, PivotColumns, PivotData, IncludePivotChart and ChartType values.
        .PARAMETER ConditionalFormat
            One or more conditional formatting rules defined with New-ConditionalFormattingIconSet.
        .PARAMETER ConditionalText
            Applies a 'Conditional formatting rule' in Excel on all the cells. When specific conditions are met a rule is triggered.
          .PARAMETER BoldTopRow
            Makes the top Row boldface.
        .PARAMETER NoHeader
            Does not put field names at the top of columns.
        .PARAMETER RangeName
            Makes the data in the worksheet a named range.
        .PARAMETER AutoNameRange
            Makes each column a named range.
        .PARAMETER TableName
            Makes the data in the worksheet a table with a name applies a style to it. Name must not contain spaces.
        .PARAMETER TableStyle
            Selects the style for the named table - defaults to 'Medium6'.
        .PARAMETER BarChart
            Creates a "quick" bar chart using the first text column as labels and the first numeric column as values
        .PARAMETER ColumnChart
            Creates a "quick" column chart using the first text column as labels and the first numeric column as values
        .PARAMETER LineChart
            Creates a "quick" line chart using the first text column as labels and the first numeric column as values
        .PARAMETER PieChart
            Creates a "quick" pie chart using the first text column as labels and the first numeric column as values
        .PARAMETER ExcelChartDefinition
            A hash table containing ChartType, Title, NoLegend, ShowCategory, ShowPercent, Yrange, Xrange and SeriesHeader for one or more [non-pivot] charts.
        .PARAMETER StartRow
            Row to start adding data. 1 by default. Row 1 will contain the title if any. Then headers will appear (Unless -No header is specified) then the data appears.
        .PARAMETER StartColumn
            Column to start adding data - 1 by default.
        .PARAMETER FreezeTopRow
            Freezes headers etc. in the top row.
        .PARAMETER FreezeFirstColumn
            Freezes titles etc. in the left column.
        .PARAMETER FreezeTopRowFirstColumn
             Freezes top row and left column (equivalent to Freeze pane 2,2 ).
        .PARAMETER FreezePane
             Freezes panes at specified coordinates (in the form  RowNumber , ColumnNumber).
        .PARAMETER AutoFilter
            Enables the 'Filter' in Excel on the complete header row. So users can easily sort, filter and/or search the data in the select column from within Excel.
        .PARAMETER AutoSize
            Sizes the width of the Excel column to the maximum width needed to display all the containing data in that cell.
        .PARAMETER Show
            Opens the Excel file immediately after creation. Convenient for viewing the results instantly without having to search for the file first.
        .PARAMETER CellStyleSB
            A script block which is run at the end of the process to apply styles to cells (although it can be used for other purposes).
            The script block is given three paramaters; an object containing the current worksheet, the Total number of Rows and the number of the last column.
        .PARAMETER ReturnRange
            If specified, Export-Excel returns the range of added cells in the format "A1:Z100"
        .PARAMETER PassThru
            If specified, Export-Excel returns an object representing the Excel package without saving the package first. To save it you need to call the save or Saveas method or send it back to Export-Excel.

      .EXAMPLE
        C:\> Send-SQLDataToExcel -MsSQLserver -Connection localhost -SQL  "select name,type,type_desc from [master].[sys].[all_objects]" -Path .\temp.xlsx -WorkSheetname master -AutoSize -FreezeTopRow -AutoFilter -BoldTopRow

        Connects to the local SQL server and selects 3 columns from [Sys].[all_objects] and exports then to a sheet named master with some basic header management
      .EXAMPLE
        C:\> $SQL="SELECT top 25 Name,Length  From TestData ORDER BY Length DESC"
        C:\> $Connection = ' Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=C:\Users\James\Documents\Database1.accdb;'

        C:\> Send-SQLDataToExcel -Connection  $connection -SQL $sql -path .\demo1.xlsx -WorkSheetname "Sizes" -AutoSize

         This declares a SQL statement and creates an  ODBC connection string to read from an Access file and extracts data from it and sends it to a new worksheet

      .EXAMPLE
        C:\> $SQL="SELECT top 25 DriverName, Count(RaceDate) as Races, Count(Win) as Wins, Count(Pole) as Poles, Count(FastestLap) as Fastlaps FROM Results GROUP BY DriverName ORDER BY (count(win)) DESC"
        C:\> $Connection = 'Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=C:\users\James\Documents\f1Results.xlsx;'

        C:\> Send-SQLDataToExcel -Connection  $connection -SQL $sql -path .\demo1.xlsx -WorkSheetname "Winners" -AutoSize -AutoNameRange -ConditionalFormat @{DataBarColor="Blue"; Range="Wins"}

        This declares a SQL statement and creates an  ODBC connection string to read from an Excel file, it then runs the statement and outputs the resulting data to a new spreadsheet.
        The spreadsheet is formatted and a data bar added to show make the drivers' wins clearer.
        (the F1 results database is available from https://1drv.ms/x/s!AhfYu7-CJv4ehNdZWxJE9LMAX_N5sg )
      .EXAMPLE
        C:\> $SQL = "SELECT top 25 DriverName, Count(RaceDate) as Races, Count(Win) as Wins, Count(Pole) as Poles, Count(FastestLap) as Fastlaps FROM Results GROUP BY DriverName ORDER BY (count(win)) DESC"
        C:\> Get-SQL -Session F1 -excel -Connection "C:\Users\mcp\OneDrive\public\f1\f1Results.xlsx" -sql $sql -OutputVariable Table | out-null

        C:\> Send-SQLDataToExcel -DataTable $Table -Path ".\demo3.xlsx" -WorkSheetname Gpwinners -autosize  -TableName winners -TableStyle Light6 -show

        This uses Get-SQL (at least V1.1 - download from the gallery with Install-Module -Name GetSQL - note the function is Get-SQL the module is GetSQL without the "-" )
        to simplify making database connections and building /submitting SQL statements.
        Here it uses the same SQL statement as before; -OutputVariable leaves a System.Data.DataTable object in $table
        and Send-SQLDataToExcel puts $table into the worksheet and sets it as an Excel table.
        (the F1 results database is available from https://1drv.ms/x/s!AhfYu7-CJv4ehNdZWxJE9LMAX_N5sg )
      .EXAMPLE
        C:\> $SQL = "SELECT top 25 DriverName,  Count(Win) as Wins FROM Results GROUP BY DriverName ORDER BY (count(win)) DESC"
        C:\> Send-SQLDataToExcel -Session $DbSessions["f1"] -SQL $sql -Path  ".\demo3.xlsx" -WorkSheetname Gpwinners -autosize -ColumnChart

        Like the previous example, this uses Get-SQL (download from the gallery with Install-Module -Name GetSQL). It uses the connection which Get-SQL made rather than an ODFBC connection string
        Here the data is presented as a quick chart.
      .EXAMPLE
        C:\>  Send-SQLDataToExcel -path .\demo3.xlsx -WorkSheetname "LR" -Connection "DSN=LR" -sql "SELECT name AS CollectionName FROM AgLibraryCollection Collection ORDER BY CollectionName"

        This example uses an Existing ODBC datasource name "LR" which maps to an adobe lightroom database and gets a list of collection names into a worksheet
    #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword","")]
     param (
        [Parameter(ParameterSetName="SQLConnection", Mandatory=$true)]
        [Parameter(ParameterSetName="ODBCConnection",Mandatory=$true)]
        $Connection,
        [Parameter(ParameterSetName="ExistingSession",Mandatory=$true)]
        [System.Data.Common.DbConnection]$Session,
        [Parameter(ParameterSetName="SQLConnection",Mandatory=$true)]
        [switch]$MsSQLserver,
        [Parameter(ParameterSetName="SQLConnection")]
        [String]$DataBase,
        [Parameter(ParameterSetName="SQLConnection", Mandatory=$true)]
        [Parameter(ParameterSetName="ODBCConnection",Mandatory=$true)]
        [Parameter(ParameterSetName="ExistingSession",Mandatory=$true)]
        [string]$SQL,
        [int]$QueryTimeout,
        [Parameter(ParameterSetName="Pre-FetchedData",Mandatory=$true)]
        [System.Data.DataTable]$DataTable,
        $Path,
        [String]$WorkSheetname = 'Sheet1',
        [Switch]$KillExcel,
        [Switch]$Show,
        [String]$Title,
        [OfficeOpenXml.Style.ExcelFillStyle]$TitleFillPattern = 'None',
        [Switch]$TitleBold,
        [Int]$TitleSize = 22,
        $TitleBackgroundColor,
        [String]$Password,
        [Hashtable]$PivotTableDefinition,
        [Switch]$IncludePivotTable,
        [String[]]$PivotRows,
        [String[]]$PivotColumns,
        $PivotData,
        [String[]]$PivotFilter,
        [Switch]$PivotDataToColumn,
        [Switch]$NoTotalsInPivot,
        [Switch]$IncludePivotChart,
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = 'Pie',
        [Switch]$NoLegend,
        [Switch]$ShowCategory,
        [Switch]$ShowPercent,
        [Switch]$AutoSize,
        [Switch]$FreezeTopRow,
        [Switch]$FreezeFirstColumn,
        [Switch]$FreezeTopRowFirstColumn,
        [Int[]]$FreezePane,
        [Switch]$AutoFilter,
        [Switch]$BoldTopRow,
        [Switch]$NoHeader,
        [String]$RangeName,
        [String]$TableName,
        [OfficeOpenXml.Table.TableStyles]$TableStyle = 'Medium6',
        [Switch]$Barchart,
        [Switch]$PieChart,
        [Switch]$LineChart ,
        [Switch]$ColumnChart ,
        [Object[]]$ExcelChartDefinition,
        [Switch]$AutoNameRange,
        [Object[]]$ConditionalFormat,
        [Object[]]$ConditionalText,
        [ScriptBlock]$CellStyleSB,
        [Int]$StartRow    = 1,
        [Int]$StartColumn = 1,
        [Switch]$ReturnRange,
        [Switch]$Passthru
    )

    if ($KillExcel) {
            Get-Process excel -ErrorAction Ignore | Stop-Process
            while (Get-Process excel -ErrorAction Ignore) {Start-Sleep -Milliseconds 250}
    }

    #We were either given a session object or a connection string (with, optionally a MSSQLServer parameter)
    # If we got -MSSQLServer, create a SQL connection, if we didn't but we got -Connection create an ODBC connection
    if     ($MsSQLserver -and $Connection) {
            if ($Connection -notmatch "=") {$Connection = "server=$Connection;trusted_connection=true;timeout=60"}
            $Session     = New-Object -TypeName System.Data.SqlClient.SqlConnection  -ArgumentList $Connection
            if ($Session.State -ne 'Open') {$Session.Open()}
            if ($DataBase) {$Session.ChangeDatabase($DataBase) }
    }
    elseif ($Connection)  {
            $Session     = New-Object -TypeName System.Data.Odbc.OdbcConnection      -ArgumentList $Connection ; $Session.ConnectionTimeout = 30
    }

    If ($session) {
        #A session was either passed in or just created. If it's a SQL one make a SQL DataAdapter, otherwise make an ODBC one
        if ($Session.GetType().name -match "SqlConnection") {
            $dataAdapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter -ArgumentList (
                           New-Object -TypeName System.Data.SqlClient.SqlCommand     -ArgumentList $SQL, $Session)
        }
        else {
            $dataAdapter = New-Object -TypeName System.Data.Odbc.OdbcDataAdapter     -ArgumentList (
                           New-Object -TypeName System.Data.Odbc.OdbcCommand         -ArgumentList $SQL, $Session )
        }
        if ($QueryTimeout) {$dataAdapter.SelectCommand.CommandTimeout = $ServerTimeout}

        #Both adapter types output the same kind of table, create one and fill it from the adapter
        $DataTable       = New-Object -TypeName System.Data.DataTable
        $rowCount        = $dataAdapter.fill($dataTable)
        Write-Verbose -Message "Query returned $rowCount row(s)"
    }
    if ($DataTable.Rows.Count) {
        #ExportExcel user a -NoHeader parameter so that's what we use here, but needs to be the other way around.
        $printHeaders    = -not $NoHeader
        if ($Title)  {$r = $StartRow +1 }
        else         {$r = $StartRow}
        #Get our Excel sheet and fill it with the data
        $excelPackage    = Export-Excel -Path $Path -WorkSheetname $WorkSheetname  -PassThru
        $excelPackage.Workbook.Worksheets[$WorkSheetname].Cells[$r,$StartColumn].LoadFromDataTable($dataTable, $printHeaders )  | Out-Null

        #Apply date format
        for ($c=0 ; $c -lt $DataTable.Columns.Count ; $c++) {
            if ($DataTable.Columns[$c].DataType -eq [datetime]) {
                Set-ExcelColumn -Worksheet $excelPackage.Workbook.Worksheets[$WorkSheetname] -Column ($c +1) -NumberFormat 'Date-Time'
            }
        }

        #Call export-excel with any parameters which don't relate to the SQL query
        "Connection", "Database" , "Session", "MsSQLserver", "Destination" , "SQL" , "DataTable", "Path" | ForEach-Object {$null = $PSBoundParameters.Remove($_) }
        Export-Excel -ExcelPackage $excelPackage   @PSBoundParameters
    }
    else {Write-Warning -Message "No Data to insert."}
    #If we were passed a connection and opened a session,  close that session.
    if ($Connection)  {$Session.close() }
}
