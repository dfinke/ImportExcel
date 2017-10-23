function ConvertFrom-ExcelToSQLInsert {
    <#
  
    .PARAMETER ColumnMap
    Hashtable of input to output column names.
    
    .PARAMETER Unique
    Suppress duplicate lines from output.
    #>
    param(
        [Parameter(Mandatory = $true)]
        $TableName,
        [Alias("FullName")]
        [Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true, Mandatory = $true)]
        [ValidateScript( { Test-Path $_ -PathType Leaf })]
        $Path,
        [Alias("Sheet")]
        $WorkSheetname = 1,
        $ColumnMap,
        [switch]$Unique,
        [int]$HeaderRow = 1,
        [string[]]$Header,
        [switch]$NoHeader,
        [switch]$DataOnly
    )


    $null = $PSBoundParameters.Remove('TableName')
    $null = $PSBoundParameters.Remove('ColumnMap')
    $null = $PSBoundParameters.Remove('Unique')
    $params = @{} + $PSBoundParameters

    $script:raw_output = @()
    ConvertFrom-ExcelData @params {
        param($propertyNames, $record)

        $iterator = @()
        If ($ColumnMap.Count -gt 0)
            {
                $iterator = $ColumnMap.Keys
                $ColumnNames = "'" + ($ColumnMap.Values -join "', '") + "'"
            }
        Else
            {
                $iterator = $PropertyNames
                $ColumnNames = "'" + ($propertyNames -join "', '") + "'"                
            }

        $values = foreach ($propertyName in $iterator) { $record.$propertyName }        
        $targetValues = "'" + ($values -join "', '") + "'"

        If ( $Unique -eq $true)
        {
            $script:raw_output += $("INSERT INTO {0} ({1}) Values({2});" -f $TableName, $ColumnNames, $targetValues)
        } Else
        {
            "INSERT INTO {0} ({1}) Values({2});" -f $TableName, $ColumnNames, $targetValues
        }
    }

    If ($Unique) { $script:raw_output | Sort-Object | Get-Unique }

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
