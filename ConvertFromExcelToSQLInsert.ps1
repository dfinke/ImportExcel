function ConvertFrom-ExcelToSQLInsert {
    <#
    .SYNOPSIS
        Generate SQL insert statements from Excel spreadsheet.

    .DESCRIPTION
        Generate SQL insert statements from Excel spreadsheet.

    .PARAMETER TableName
        Name of the target database table.

    .PARAMETER Path
        Path to an existing .XLSX file

        This parameter is passed to Import-Excel as is.

    .PARAMETER WorkSheetname
        Specifies the name of the worksheet in the Excel workbook to import. By default, if no name is provided, the first worksheet will be imported.

        This parameter is passed to Import-Excel as is.

    .PARAMETER StartRow
        The row from where we start to import data, all rows above the StartRow are disregarded. By default this is the first row.

        When the parameters ‘-NoHeader’ and ‘-HeaderName’ are not provided, this row will contain the column headers that will be used as property names. When one of both parameters are provided, the property names are automatically created and this row will be treated as a regular row containing data.

    .PARAMETER Header
        Specifies custom property names to use, instead of the values defined in the column headers of the TopRow.

        In case you provide less header names than there is data in the worksheet, then only the data with a corresponding header name will be imported and the data without header name will be disregarded.

        In case you provide more header names than there is data in the worksheet, then all data will be imported and all objects will have all the property names you defined in the header names. As such, the last properties will be blanc as there is no data for them.

    .PARAMETER NoHeader
        Automatically generate property names (P1, P2, P3, ..) instead of the ones defined in the column headers of the TopRow.

        This switch is best used when you want to import the complete worksheet ‘as is’ and are not concerned with the property names.

    .PARAMETER DataOnly
        Import only rows and columns that contain data, empty rows and empty columns are not imported.


    .PARAMETER ConvertEmptyStringsToNull
        If specified, cells without any data are replaced with NULL, instead of an empty string.

        This is to address behviors in certain DBMS where an empty string is insert as 0 for INT column, instead of a NULL value.

    .EXAMPLE
        Generate SQL insert statements from Movies.xlsx file, leaving blank cells as empty strings:

        ----------------------------------------------------------
        | File: Movies.xlsx            -           Sheet: Sheet1 |
        ----------------------------------------------------------
        |           A            B            C                  |
        |1     Movie Name      Year         Rating               |
        |2     The Bodyguard   1992           9                  |
        |3     The Matrix      1999           8                  |
        |4     Skyfall         2012           9                  |
        |5     The Avengers    2012                              |
        ----------------------------------------------------------

        PS C:\> Import-Excel -TableName "Movies" -Path 'C:\Movies.xlsx'
        INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('The Bodyguard', '1992', '9');
        INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('The Matrix', '1999', '8');
        INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('Skyfall', '2012', '9');
        INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('The Avengers', '2012', '');

    .EXAMPLE
        Generate SQL insert statements from Movies.xlsx file, specify NULL instead of an empty string.

        ----------------------------------------------------------
        | File: Movies.xlsx            -           Sheet: Sheet1 |
        ----------------------------------------------------------
        |           A            B            C                  |
        |1     Movie Name      Year         Rating               |
        |2     The Bodyguard   1992           9                  |
        |3     The Matrix      1999           8                  |
        |4     Skyfall         2012           9                  |
        |5     The Avengers    2012                              |
        ----------------------------------------------------------

        PS C:\> ConvertFrom-ExcelToSQLInsert -TableName "Movies" -Path "C:\Movies.xlsx" -ConvertEmptyStringsToNull
        INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('The Bodyguard', '1992', '9');
        INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('The Matrix', '1999', '8');
        INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('Skyfall', '2012', '9');
        INSERT INTO Movies ('Movie Name', 'Year', 'Rating') Values('The Avengers', '2012', NULL);

    .NOTES
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $TableName,
        [Alias("FullName")]
        [Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true, Mandatory = $true)]
        [ValidateScript( { Test-Path $_ -PathType Leaf })]
        $Path,
        [Alias("Sheet")]
        $WorkSheetname = 1,
        [Alias('HeaderRow', 'TopRow')]
        [ValidateRange(1, 9999)]
        [Int]$StartRow,
        [string[]]$Header,
        [switch]$NoHeader,
        [switch]$DataOnly,
        [switch]$ConvertEmptyStringsToNull
    )

    $null = $PSBoundParameters.Remove('TableName')
    $null = $PSBoundParameters.Remove('ConvertEmptyStringsToNull')

    $params = @{} + $PSBoundParameters

    ConvertFrom-ExcelData @params {
        param($propertyNames, $record)

        $ColumnNames = "'" + ($PropertyNames -join "', '") + "'"
        $values = foreach ($propertyName in $PropertyNames) {
            if ($ConvertEmptyStringsToNull.IsPresent -and [string]::IsNullOrEmpty($record.$propertyName)) {
                'NULL'
            }
            else {
                "'" + $record.$propertyName + "'"
            }
        }
        $targetValues = ($values -join ", ")

        "INSERT INTO {0} ({1}) Values({2});" -f $TableName, $ColumnNames, $targetValues
    }
}