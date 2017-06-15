function ConvertFrom-ExcelData {
    <#
    .SYNOPSIS
    Reads data from a sheet, and for each row, calls a custom scriptblock with a list of property names and the row of data.

    
    .EXAMPLE
    ConvertFrom-ExcelData .\testSQLGen.xlsx {
        param($propertyNames, $record)

        $reportRecord = @()
        foreach ($pn in $propertyNames) {
            $reportRecord += "{0}: {1}" -f $pn, $record.$pn
        }
        $reportRecord +=""
        $reportRecord -join "`r`n"
}

First: John
Last: Doe
The Zip: 12345
....
    #>
    param(
        [Alias("FullName")]
        [Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true, Mandatory = $true)]
        [ValidateScript( { Test-Path $_ -PathType Leaf })]
        $Path,
        [ScriptBlock]$scriptBlock,
        [Alias("Sheet")]
        $WorkSheetname = 1,
        [int]$HeaderRow = 1,
        [string[]]$Header,
        [switch]$NoHeader,
        [switch]$DataOnly
    )

    $null = $PSBoundParameters.Remove('scriptBlock')
    $params = @{} + $PSBoundParameters

    $data = Import-Excel @params

    $PropertyNames = $data[0].psobject.Properties |
        Where-Object {$_.membertype -match 'property'} |
        Select-Object -ExpandProperty name

    foreach ($record in $data) {
        & $scriptBlock $PropertyNames $record
    }
}