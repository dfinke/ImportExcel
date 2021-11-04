#Requires -Version 5
function Invoke-ExcelQuery {
    <#
        .SYNOPSIS
        Helper method for executing Read-OleDbData with some basic defaults.

        For additional help, see documentation for Read-OleDbData cmdlet.

        .DESCRIPTION
        Uses Read-OleDbData to execute a sql statement against a xlsx file. For finer grained control over the interaction, you may use that cmdlet. This cmdlet assumes a file path will be passed in and the connection string will be built with no headers and treating all results as text.

        Running this command is equivalent to running the following:

        $FullName = (Get-ChildItem $Path).FullName
        Read-OleDbData `
            -ConnectionString "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$FullName;Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'" `
            -SqlStatement $Query

        .EXAMPLE
        Invoke-ExcelQuery .\test.xlsx 'select ROUND(F1) as [A1] from [sheet3$A1:A1]'

        .EXAMPLE
        $Path = (Get-ChildItem 'test.xlsx').FullName
        $Query = "select ROUND(F1) as [A] from [sheet1$A1:A1]"
        Read-XlsxUsingOleDb -Path $Path -Query $Query

        .EXAMPLE
        $ReadDataArgs = @{
            Path = .\test.xlsx
            Query = Get-Content query.sql -Raw
        }
        $Results = Invoke-ExcelQuery @ReadDataArgs
    #>
    param(
        #The path to the file to open.
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String] $Path, # var name consistent with Import-Excel
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String] $Query # var name consistent with Invoke-Sqlcmd
    )
    $FullName = (Get-ChildItem $Path).FullName
    Invoke-ExcelQuery `
        -ConnectionString "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$FullName;Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'" `
        -SqlStatement $Query
}