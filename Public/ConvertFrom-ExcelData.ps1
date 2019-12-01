function ConvertFrom-ExcelData {
    [alias("Use-ExcelData")]
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