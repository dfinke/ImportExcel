function ConvertFrom-ExcelToSQLInsert {
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
        [switch]$ConvertEmptyStringsToNull,
        [switch]$UseMSSQLSyntax
    )

    $null = $PSBoundParameters.Remove('TableName')
    $null = $PSBoundParameters.Remove('ConvertEmptyStringsToNull')
    $null = $PSBoundParameters.Remove('UseMSSQLSyntax')

    $params = @{} + $PSBoundParameters

    ConvertFrom-ExcelData @params {
        param($propertyNames, $record)

        $ColumnNames = "'" + ($PropertyNames -join "', '") + "'"
        if($UseMSSQLSyntax) {
            $ColumnNames = "[" + ($PropertyNames -join "], [") + "]"
        }

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