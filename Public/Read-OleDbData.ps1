#Requires -Version 5
function Read-OleDbData {
    <#
        .SYNOPSIS
        Read data from an OleDb source using dotnet classes. This allows for OleDb queries against excel spreadsheets. Examples will only be for querying xlsx files.

        For additional documentation, see Microsoft's documentation on the System.Data OleDb namespace here:
        https://docs.microsoft.com/en-us/dotnet/api/system.data.oledb

        .DESCRIPTION
        Read data from an OleDb source using dotnet classes. This allows for OleDb queries against excel spreadsheets. Examples will only be for querying xlsx files using ACE.

        .EXAMPLE
        Read-OleDbData `
            -ConnectionString "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=file.xlsx;Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'" `
            -SqlStatement "select ROUND(F1) as [A] from [sheet1$A1:A1]"

        .EXAMPLE
        $ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=file.xlsx;Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'"
        $SqlStatement = "select ROUND(F1) as [A] from [sheet1$A1:A1]"
        Read-OleDbData -ConnectionString $ConnectionString -SqlStatement $SqlStatement

        .EXAMPLE
        $ReadDataArgs = @{
            SqlStatement = Get-Content query.sql -Raw
            ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=file.xlsx;Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'"
        }
        $Results = Read-OleDbData @ReadDataArgs
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String] $ConnectionString,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String] $SqlStatement   
    )
    
    if ($IsLinux -or $IsMacOS) {
        #todo: possibly add support for linux/mac somehow. i haven't researched this at all as i 
        #      don't have a need for that and i'm not sure anyone else would in this context, but it does
        #      appear that once upon a time mono had support for oledb, so maybe it is (or was) supported.
        #      mono link here: https://www.mono-project.com/archived/ole_db/ 
        Write-Error "Read-OleDbData only runs on Windows"
        return
    }

    #todo: add checks for dotnet libs
    #todo: possibly add checks for ace drivers, but maybe only needed if we want to restrict usage.
    #      i currently just pass through the query and connection string so user is only limited by
    #      their own machine setup, but they have to check for those dependencies themselves.
    #todo: possibly try/catch. i personally do not do this, as i let them throw naturally and catch
    #      them/handle them outside of this function.
    #todo: possibly allow for DataSets instead of just DataTable. I used to use a similar method before
    #      switching to the sqlcmd module and use datasets as I had multiple tables coming back from
    #      sql sometimes. i think in this case, it's best to just keep it simple, but maybe someone
    #      out there would prefer to be able to get multiple tables back. i have not tested that
    #      with the OleDbDataAdapter.
    #todo: possibly just return the datatable, i do it as below because i prefer to simplify the output
    #      and get rid of the extra fields that come back with the datatable or rows.
    
    $DataTable = new-object System.Data.DataTable
    $DataAdapter = new-object System.Data.OleDb.OleDbDataAdapter $SqlStatement, $ConnectionString
    $null = $DataAdapter.Fill($DataTable)
    $null = $DataAdapter.Dispose()
    $DataTable.Rows | Select-Object $DataTable.Columns.ColumnName
}