function ConvertFrom-ExcelToSQLInsert {
    param(
        [Parameter(Mandatory = $true)]
        $TableName,
        [Alias("FullName")]
        [Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true, Mandatory = $true)]
        [ValidateScript( { Test-Path $_ -PathType Leaf })]
        $Path,
        [Alias("Sheet")]
        $WorkSheetname = 1,
        [int]$HeaderRow = 1,
        [string[]]$Header,
        [switch]$NoHeader,
        [switch]$DataOnly
    )

    $null = $PSBoundParameters.Remove('TableName')
    $params = @{} + $PSBoundParameters

    ConvertFrom-ExcelData @params {
        param($propertyNames, $record)

        $ColumnNames = "'" + ($PropertyNames -join "', '") + "'"
        $values = foreach ($propertyName in $PropertyNames) { $record.$propertyName }
        $targetValues = "'" + ($values -join "', '") + "'"

        "INSERT INTO {0} ({1}) Values({2});" -f $TableName, $ColumnNames, $targetValues
    }
    # $data = Import-Excel @params    
    
    # $PropertyNames = $data[0].psobject.Properties |
    #     Where-Object {$_.membertype -match 'property'} |
    #     Select-Object -ExpandProperty name
    
    # $ColumnNames = "'" + ($PropertyNames -join "', '") + "'"

    # foreach ($record in $data) {
    #     $values = $(foreach ($propertyName in $PropertyNames) {
    #             $record.$propertyName
    #         })

    #     $targetValues = "'" + ($values -join "', '") + "'"

    #     "INSERT INTO {0} ({1}) Values({2});" -f $TableName, $ColumnNames, $targetValues
    # }
}