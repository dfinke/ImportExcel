function Send-SQLDataToExcel {
    [CmdletBinding(DefaultParameterSetName="none")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification="Allowed to use DBSessions Global variable from GETSQL Module")]

    param (
        [Parameter(ParameterSetName="SQLConnection",   Mandatory=$true)]
        [Parameter(ParameterSetName="ODBCConnection",  Mandatory=$true)]
        $Connection,
        [Parameter(ParameterSetName="ExistingSession", Mandatory=$true)]
        $Session,
        [Parameter(ParameterSetName="SQLConnection",   Mandatory=$true)]
        [switch]$MsSqlServer,
        [Parameter(ParameterSetName="SQLConnection")]
        [String]$DataBase,
        [Parameter(ParameterSetName="SQLConnection",   Mandatory=$true)]
        [Parameter(ParameterSetName="ODBCConnection",  Mandatory=$true)]
        [Parameter(ParameterSetName="ExistingSession", Mandatory=$true)]
        [string]$SQL,
        [int]$QueryTimeout,
        [Parameter(ParameterSetName="Pre-FetchedData", Mandatory=$true)]
        [System.Data.DataTable]$DataTable,
        [switch]$Force
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
      #region Dynamic params mean we can get passed parameter combination Export-Excel will reject, so throw here, rather than get data and then have Export-Excel error.
        if ($PSBoundParameters.Path -and $PSBoundParameters.ExcelPackage) {
            throw 'Parameter error: you cannot specify both a path and an Excel Package.'
            return
        }
        if ($PSBoundParameters.AutoFilter -and ($PSBoundParameters.TableName -or $PSBoundParameters.TableStyle)) {
            Write-Warning "Tables are automatically auto-filtered, -AutoFilter will be ignored"
            $null = $PSBoundParameters.Remove('AutoFilter')
        }
      #endregion
      #region if we were either given a session object or a connection string (& optionally -MsSqlServer) make sure we can connect
        try {
            #If we got -MsSqlServer, create a SQL connection, if we didn't but we got -Connection create an ODBC connection
            if     ($MsSqlServer -and $Connection) {
                if ($Connection -notmatch '=') {$Connection = "server=$Connection;trusted_connection=true;timeout=60"}
                $Session     = New-Object -TypeName System.Data.SqlClient.SqlConnection  -ArgumentList $Connection
                if ($Session.State -ne 'Open') {$Session.Open()}
                if ($DataBase) {$Session.ChangeDatabase($DataBase) }
            }
            elseif ($Connection)  {
                $Session     = New-Object -TypeName System.Data.Odbc.OdbcConnection      -ArgumentList $Connection ; $Session.ConnectionTimeout = 30
            }
        }
        catch {
            Write-Warning "An Error occured trying to connect to $Connection, the error was $([Environment]::NewLine + $_.Exception.InnerException))"
        }
        if ($Session -is [String] -and $Global:DbSessions[$Session]) {$Session = $Global:DbSessions[$Session]}
      #endregion
      #region we may have been given a table, but if there is a db session to connect to, send the query
        if     ($Session) {
            try {
                #If the session a SQL one make a SQL DataAdapter, otherwise make an ODBC one
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
            catch {
                Write-Warning "An Error occured trying to run the query, the error was $([Environment]::NewLine + $_.Exception.InnerException))"
            }
        }
      #endregion
      #region send the table to Excel
        #remove parameters which relate to querying SQL, leaving the ones used by Export-Excel
        'Connection' , 'Database'  , 'Session' , 'MsSqlServer' , 'SQL'  , 'DataTable'  , 'QueryTimeout' , 'Force' |
                ForEach-Object {$null = $PSBoundParameters.Remove($_) }
        #if force was specified export even if there are no rows. If there are no columns, the query failed and export "null" if forced
        if     ($DataTable.Rows.Count) {
             Export-Excel  @PSBoundParameters -InputObject $DataTable
        }
        elseif ($Force -and $DataTable.Columns.Count) {
            Write-Warning -Message "Zero rows returned, and -Force was specified, sending empty table to Excel."
            Export-Excel  @PSBoundParameters -InputObject $DataTable
        }
        elseif ($Force) {
            Write-Warning -Message "-Force was specified but there is no data to send."
            Export-Excel  @PSBoundParameters -InputObject $null
        }
        else   {Write-Warning -Message 'There is no Data to insert, and -Force was not specified.' }
      #endregion
      #If we were passed a connection and opened a session,  close that session.
        if     ($Connection)  {$Session.close() }
    }
}