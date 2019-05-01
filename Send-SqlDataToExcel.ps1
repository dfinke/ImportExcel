Function Send-SQLDataToExcel {
        <#
      .SYNOPSIS
        Inserts a DataTable - returned by a SQL query - into an ExcelSheet
      .DESCRIPTION
        This command takes a SQL statement and run it against a database connection; for the connection it accepts either
        * an object representing a session with a SQL server or ODBC database, or
        * a connection string to make a session (if -MSSQLServer is specified it uses the SQL Native client,
          and -Connection can be a server name instead of a detailed connection string. Without this switch it uses ODBC)
        The command takes all the parameters of Export-Excel, except for -InputObject (alias TargetData); after
        fetching the data it calls Export-Excel with the data as the value of InputParameter and whichever of
        Export-Excel's parameters it was passed; for details of these parameters see the help for Export-Excel.
      .PARAMETER Session
        An active ODBC Connection or SQL connection object representing a session with a database which will be queried to get the data .
      .PARAMETER Connection
        A database connection string to be used to create a database session; either
        * A Data source name written in the form DSN=ODBC_Data_Source_Name, or
        * A full ODBC or SQL Native Client Connection string, or
        * The name of a SQL server.
      .PARAMETER MSSQLServer
        Specifies the connection string is for SQL server, not ODBC.
      .PARAMETER SQL
        The SQL query to run against the session which was passed in -Session or set up from -Connection.
      .PARAMETER Database
        Switches to a specific database on a SQL server.
      .PARAMETER QueryTimeout
        Override the default query time of 30 seconds.
      .PARAMETER DataTable
        A System.Data.DataTable object containing the data to be inserted into the spreadsheet without running a query.
        This remains supported to avoid breaking older scripts, but if you have a DataTable object you can pass the it
        into Export-Excel using -InputObject.
      .EXAMPLE
        C:\> Send-SQLDataToExcel -MsSQLserver -Connection localhost -SQL  "select name,type,type_desc from [master].[sys].[all_objects]" -Path .\temp.xlsx -WorkSheetname master -AutoSize -FreezeTopRow -AutoFilter -BoldTopRow

        Connects to the local SQL server and selects 3 columns from [Sys].[all_objects] and exports then to a sheet named master with some basic header management
      .EXAMPLE
        C:\> $dbPath = 'C:\Users\James\Documents\Database1.accdb'
        C:\> $Connection = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=$dbPath;"
        C:\> $SQL="SELECT top 25 Name,Length  From TestData ORDER BY Length DESC"

        C:\> Send-SQLDataToExcel -Connection  $connection -SQL $sql -path .\demo1.xlsx -WorkSheetname "Sizes" -AutoSize

        This creates an  ODBC connection string to read from an Access file and a SQL Statement to extracts data from it,
        and sends the resulting data to a new worksheet

      .EXAMPLE
        C:\> $dbPath = 'C:\users\James\Documents\f1Results.xlsx'
        C:\> $Connection = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=$dbPath;"
        C:\> $SQL="SELECT top 25 DriverName, Count(RaceDate) as Races, Count(Win) as Wins, Count(Pole) as Poles, Count(FastestLap) as Fastlaps " +
                   " FROM Results  GROUP BY DriverName ORDER BY (count(win)) DESC"

        C:\> Send-SQLDataToExcel -Connection  $connection -SQL $sql -path .\demo2.xlsx -WorkSheetname "Winners" -AutoSize -AutoNameRange -ConditionalFormat @{DataBarColor="Blue"; Range="Wins"}

        Similar to the previous example this creates a connection string, this time for an Excel file, and runs
        a SQL statement to get a list of motor-racing results, outputting the resulting data to a new spreadsheet.
        The spreadsheet is formatted and a data bar added to show make the drivers' wins clearer.
        (the F1 results database is available from https://1drv.ms/x/s!AhfYu7-CJv4ehNdZWxJE9LMAX_N5sg )
      .EXAMPLE
        C:\> $dbPath = 'C:\users\James\Documents\f1Results.xlsx'
        C:\> $SQL = "SELECT top 25 DriverName, Count(RaceDate) as Races, Count(Win) as Wins, Count(Pole) as Poles, Count(FastestLap) as Fastlaps " +
                     " FROM Results GROUP BY DriverName ORDER BY (count(win)) DESC"
        C:\> $null = Get-SQL -Session F1 -excel -Connection $dbPath -sql $sql -OutputVariable Table

        C:\> Send-SQLDataToExcel -DataTable $Table -Path ".\demo3.xlsx" -WorkSheetname Gpwinners -autosize  -TableName winners -TableStyle Light6 -show

        This uses Get-SQL (at least V1.1 - download from the PowerShell gallery with Install-Module -Name GetSQL -
        note the function is Get-SQL the module is GetSQL without the "-" )
        Get-SQL simplify making database connections and building /submitting SQL statements.
        Here Get-SQL uses the same SQL statement as before; -OutputVariable leaves a System.Data.DataTable object in $table
        and Send-SQLDataToExcel puts $table into the worksheet and sets it as an Excel table.
        The command is equivalent to running
        C:\>  Export-Excel -inputObject $Table -Path ".\demo3.xlsx" -WorkSheetname Gpwinners -autosize  -TableName winners -TableStyle Light6 -show
        This is quicker than using
        C:\> Get-SQL <parameters> | export-excel -ExcludeProperty rowerror,rowstate,table,itemarray,haserrors <parameters>
        (the F1 results database is available from https://1drv.ms/x/s!AhfYu7-CJv4ehNdZWxJE9LMAX_N5sg )
      .EXAMPLE
        C:\> $SQL = "SELECT top 25 DriverName,  Count(Win) as Wins FROM Results GROUP BY DriverName ORDER BY (count(win)) DESC"
        C:\> Send-SQLDataToExcel -Session $DbSessions["f1"] -SQL $sql -Path  ".\demo3.xlsx" -WorkSheetname Gpwinners  -ClearSheet -autosize -ColumnChart

        Like the previous example, this uses Get-SQL (download from the gallery with Install-Module -Name GetSQL).
        It uses the database session which Get-SQL created, rather than an ODBC connection string.
        The Session parameter can either be a object (as shown here), or the name used by Get-SQL ("F1" in this case).
        Here the data is presented as a quick chart.
      .EXAMPLE
        C:\>  Send-SQLDataToExcel -path .\demo4.xlsx -WorkSheetname "LR" -Connection "DSN=LR" -sql "SELECT name AS CollectionName FROM AgLibraryCollection Collection ORDER BY CollectionName"

        This example uses an Existing ODBC datasource name "LR" which maps to an adobe lightroom database and gets a list of collection names into a worksheet

    .Link
        Export-Excel
    #>
    [CmdletBinding(DefaultParameterSetName="none")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification="Allowed to use DBSessions Global variable from GETSQL Module")]

    param (
        [Parameter(ParameterSetName="SQLConnection",   Mandatory=$true)]
        [Parameter(ParameterSetName="ODBCConnection",  Mandatory=$true)]
        $Connection,
        [Parameter(ParameterSetName="ExistingSession", Mandatory=$true)]
        $Session,
        [Parameter(ParameterSetName="SQLConnection",   Mandatory=$true)]
        [switch]$MsSQLserver,
        [Parameter(ParameterSetName="SQLConnection")]
        [String]$DataBase,
        [Parameter(ParameterSetName="SQLConnection",   Mandatory=$true)]
        [Parameter(ParameterSetName="ODBCConnection",  Mandatory=$true)]
        [Parameter(ParameterSetName="ExistingSession", Mandatory=$true)]
        [string]$SQL,
        [int]$QueryTimeout,
        [Parameter(ParameterSetName="Pre-FetchedData", Mandatory=$true)]
        [System.Data.DataTable]$DataTable
    )
#Import the parameters from Export-Excel, we will pass InputObject, and we have the common parameters so exclude those,
#and re-write the [Parmameter] attribute on each one to avoid parameterSetName here competing with the settings in Export excel.
#The down side of this that impossible parameter combinations won't be filtered out and need to be caught later.
    DynamicParam {
        $ParameterAttribute  =                        "System.Management.Automation.ParameterAttribute"
        $RuntimeDefinedParam =                        "System.Management.Automation.RuntimeDefinedParameter"
        $paramDictionary     =    New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary
        $attributeCollection =    New-Object -TypeName System.Collections.ObjectModel.Collection[System.Attribute]
        $attributeCollection.Add((New-Object -TypeName $ParameterAttribute -Property @{ ParameterSetName = "__AllParameterSets" ;Mandatory = $false}))
        foreach ($P in (Get-Command -Name Export-Excel).Parameters.values.where({$_.name -notmatch 'Verbose|Debug|Action$|Variable$|Buffer$|TargetData$|InputObject$'}))  {
            $paramDictionary.Add($p.Name, (New-Object -TypeName $RuntimeDefinedParam -ArgumentList $p.name, $p.ParameterType, $attributeCollection ) )
        }
        return $paramDictionary
    }
    process {
        #Dynamic params mean we can get passed parameter combination Export-Excel will reject, so throw here, rather than get data and then have Export-Excel error.
        if ($PSBoundParameters.Path -and $PSBoundParameters.ExcelPackage) {
            throw 'Parameter error: you cannot specify both a path and an Excel Package.'
            return
        }
        if ($PSBoundParameters.AutoFilter -and ($PSBoundParameters.TableName -or $PSBoundParameters.TableStyle)) {
            Write-Warning "Tables are automatically auto-filtered, -AutoFilter will be ignored"
            $null = $PSBoundParameters.Remove('AutoFilter')
        }
        #We were either given a session object or a connection string (with, optionally a MSSQLServer parameter)
        #If we got -MSSQLServer, create a SQL connection, if we didn't but we got -Connection create an ODBC connection
        if     ($MsSQLserver -and $Connection) {
                if ($Connection -notmatch '=') {$Connection = "server=$Connection;trusted_connection=true;timeout=60"}
                $Session     = New-Object -TypeName System.Data.SqlClient.SqlConnection  -ArgumentList $Connection
                if ($Session.State -ne 'Open') {$Session.Open()}
                if ($DataBase) {$Session.ChangeDatabase($DataBase) }
        }
        elseif ($Connection)  {
                $Session     = New-Object -TypeName System.Data.Odbc.OdbcConnection      -ArgumentList $Connection ; $Session.ConnectionTimeout = 30
        }
        if     ($Session) {
            #A session was either passed in or just created. If it's a SQL one make a SQL DataAdapter, otherwise make an ODBC one
            if ($Session -is [String] -and $Global:DbSessions[$Session]) {$Session = $Global:DbSessions[$Session]}
            if ($Session.GetType().name -match "SqlConnection") {
                $dataAdapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter -ArgumentList (
                               New-Object -TypeName System.Data.SqlClient.SqlCommand     -ArgumentList $SQL, $Session)
            }
            else {
                $dataAdapter = New-Object -TypeName System.Data.Odbc.OdbcDataAdapter     -ArgumentList (
                               New-Object -TypeName System.Data.Odbc.OdbcCommand         -ArgumentList $SQL, $Session )
            }
            if ($QueryTimeout) {$dataAdapter.SelectCommand.CommandTimeout = $QueryTimeout}

            #Both adapter types output the same kind of table, create one and fill it from the adapter
            $DataTable       = New-Object -TypeName System.Data.DataTable
            $rowCount        = $dataAdapter.fill($dataTable)
            Write-Verbose -Message "Query returned $rowCount row(s)"
        }
        if     ($DataTable.Rows.Count) {
            #Call export-excel removing parameters which relate to the SQL query, and keeping the rest.
            'Connection' , 'Database'  , 'Session' , 'MsSQLserver' , 'SQL'  , 'DataTable'  , 'QueryTimeout'  | ForEach-Object {$null = $PSBoundParameters.Remove($_) }
            Export-Excel  @PSBoundParameters -InputObject $DataTable
        }
        else   {Write-Warning -Message ' No Data to insert.' }
        #If we were passed a connection and opened a session,  close that session.
        if     ($Connection)  {$Session.close() }
    }
}