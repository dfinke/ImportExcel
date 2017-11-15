Function Send-SQLDataToExcel {
<#
    .Synopsis 
        Runs a SQL query and inserts the results into an ExcelSheet, more efficiently than sending it via Export-Excel
    .Description
        This command takes either an object representing a session with a SQL server or ODBC database, or a connection String to make one.
        It the runs a SQL command, and inserts the rows of data returned into a worksheet. 
        It takes most of the parameters of Export-Excel, but it is more efficient than getting dataRows and piping them into Export-Excel,
        data-rows have additional properties which need to be stripped off. 
     .Example 
       C:\> Send-SQLDataToExcel -MsSQLserver -Connection localhost -SQL  "select name,type,type_desc from [master].[sys].[all_objects]" -Path .\temp.xlsx -WorkSheetname master -AutoSize -FreezeTopRow -AutoFilter -BoldTopRow    
       Connects to the local SQL server and selects 3 columns from [Sys].[all_objects] and exports then to a sheet named master with some basic header manager
     .Example 
        C:\> $SQL="SELECT top 25 DriverName, Count(RaceDate) as Races, Count(Win) as Wins, Count(Pole) as Poles, Count(FastestLap) as Fastlaps FROM Results GROUP BY DriverName ORDER BY (count(win)) DESC" 
        C:\> $Connection = 'Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DriverId=790;ReadOnly=0;Dbq=C:\users\James\Documents\f1Results.xlsx;' 
        C:\> Send-SQLDataToExcel -Connection  $connection -SQL $sql -path .\demo4.xlsx -WorkSheetname "Winners" -AutoSize -AutoNameRange  

        This declares a SQL statement and creates an  ODBC connection string to read from an Excel file, it then runs the statement and outputs the resulting data to a new spreadsheet.  
     .Example
        C:\>  Send-SQLDataToExcel -path .\demo4.xlsx -WorkSheetname "LR" -Connection "DSN=LR" -sql "SELECT name AS CollectionName FROM AgLibraryCollection Collection ORDER BY CollectionName" 

        This example uses an Existing ODBC datasource name "LR" which maps to an adobe lightroom database and gets a list of collection names into a worksheet 
 


#>
    param (
        #Database connection string; either DSN=ODBC_Data_Source_Name, a full odbc or SQL Connection string, or the name of a SQL server 
        [Parameter(ParameterSetName="SQLConnection", Mandatory=$true)]
        [Parameter(ParameterSetName="ODBCConnection",Mandatory=$true)]
        $Connection, 
        #A pre-existing database session object 
        [Parameter(ParameterSetName="ExistingSession",Mandatory=$true)] 
        [System.Data.Common.DbConnection]$Session,
        #Specifies the connection string is for SQL server not ODBC 
        [Parameter(ParameterSetName="SQLConnection",Mandatory=$true)]
        [switch]$MsSQLserver,
        #Switches to a specific database on a SQL server
        [Parameter(ParameterSetName="SQLConnection")]
        [switch]$DataBase,
        #The SQL query to run 
        [Parameter(Mandatory=$true)]
        [string]$SQL, 
        $Path, 
        [String]$WorkSheetname = 'Sheet1',   
        [Switch]$KillExcel, 
        #If Specified, open the file created.
        [Switch]$Show,
        [String]$Title,
        [OfficeOpenXml.Style.ExcelFillStyle]$TitleFillPattern = 'None',
        [Switch]$TitleBold,
        [Int]$TitleSize = 22,
        [System.Drawing.Color]$TitleBackgroundColor,
        [String]$Password,
        [String[]]$PivotRows,
        [String[]]$PivotColumns,
        $PivotData,
        [Switch]$PivotDataToColumn,        
        [Hashtable]$PivotTableDefinition,
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
        [Object[]]$ExcelChartDefinition,
        [Switch]$AutoNameRange,
        [Object[]]$ConditionalFormat,
        [Object[]]$ConditionalText,
        [ScriptBlock]$CellStyleSB,
        [Int]$StartRow    = 1,
        [Int]$StartColumn = 1,
        #If Specified, return an ExcelPackage object to allow further work to be done on the file. 
        [Switch]$Passthru
    )
    
    if ($KillExcel) {
            Get-Process excel -ErrorAction Ignore | Stop-Process
            while (Get-Process excel -ErrorAction Ignore) {}
    }
    
    #We were either given a session object or a connection string (with, optionally a MSSQLServer parameter)
    # If we got -MSSQLServer, create a SQL connection, if we didn't but we got -Connection create an ODBC connection
    if     ($MsSQLserver) {
            if ($connection -notmatch "=") {$Connection = "server=$Connection;trusted_connection=true;timeout=60"} 
            $Session = New-Object -TypeName System.Data.SqlClient.SqlConnection  -ArgumentList $Connection
            if ($DataBase) {$Session.ChangeDatabase($DataBase) }
    }
    elseif ($Connection)  {
            $Session = New-Object -TypeName System.Data.Odbc.OdbcConnection      -ArgumentList $Connection ; $Session.ConnectionTimeout = 30
    }

    #A session was either passed in or just created. If it's a SQL one make a SQL DataAdapter, otherwise make an ODBC one 
    if ($Session.gettype().name -match "SqlConnection") {  
        $dataAdapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter -ArgumentList (
                       New-Object -TypeName System.Data.SqlClient.SqlCommand     -ArgumentList $sql, $Session)
    }
    else {
        $dataAdapter = New-Object -TypeName System.Data.Odbc.OdbcDataAdapter     -ArgumentList (
                       New-Object -TypeName System.Data.Odbc.OdbcCommand         -ArgumentList $sql, $Session ) 
    }
    
    #Both adapter types output the same kind of table, create one and fill it from the adapter    
    $dataTable       = New-Object -TypeName System.Data.DataTable
    $rowCount        = $dataAdapter.fill($dataTable)
    Write-Verbose "Query returned $rowcount row(s)" 

    #ExportExcel user a -NoHeader parameter so that's what we use here, but needs to be the other way around. 
    $PrintHeaders    = -not $NoHeader
    if ($Title)  {$r = $StartRow +1 } 
    else         {$r = $StartRow} 
    #Get our Excel sheet and fill it with the data 
    $excelPackage    = Export-Excel -Path $Path -WorkSheetname $WorkSheetname  -PassThru
    $excelPackage.Workbook.Worksheets[$WorkSheetname].Cells[$r,$StartColumn].LoadFromDataTable($dataTable, $PrintHeaders )  | Out-Null
    
    #Call export-excel with any parameters which don't relate to the SQL query
    "Connection", "Session", "MsSQLserver", "Destination" , "sql" ,"Path" | foreach-object {$null = $PSBoundParameters.Remove($_) }
    Export-Excel -ExcelPackage $excelPackage   @PSBoundParameters 

    #If we were not passed a session close the session we created. 
    if ($Connection)  {$Session.close() } 
}
