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
    
    try {
        if ((New-Object system.data.oledb.oledbenumerator).GetElements().SOURCES_NAME -notcontains "Microsoft.ACE.OLEDB.12.0") {
            Write-Warning "Microsoft.ACE.OLEDB.12.0 provider is missing! You will not be able to query Excel files without it. Please install from https://www.microsoft.com/en-us/download/details.aspx?id=54920"
        }
    }
    catch {
        Write-Error "System.Data.OleDb is not working or you are on an unsupported platform."
        return
    }
    
    $DataTable = new-object System.Data.DataTable
    $DataAdapter = new-object System.Data.OleDb.OleDbDataAdapter $SqlStatement, $ConnectionString
    $null = $DataAdapter.Fill($DataTable)
    $null = $DataAdapter.Dispose()
    $DataTable.Rows | Select-Object $DataTable.Columns.ColumnName
}