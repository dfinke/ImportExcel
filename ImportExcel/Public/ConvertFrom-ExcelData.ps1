function ConvertFrom-ExcelData {
    [alias("Use-ExcelData")]
    param(
        [Alias("FullName")]
        [Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true, Mandatory = $true)]
        [ValidateScript( { Test-Path $_ -PathType Leaf })]
        $Path,
        [ScriptBlock]$ScriptBlock,
        [Alias("Sheet")]
        $WorksheetName = 1,
		[Alias('HeaderRow', 'TopRow')]
        [int]$StartRow = 1,
        [string[]]$Header,
        [switch]$NoHeader,
        [switch]$DataOnly
    )

    $null = $PSBoundParameters.Remove('ScriptBlock')
    $params = @{} + $PSBoundParameters

    $data = Import-Excel @params

    $PropertyNames = $data[0].psobject.Properties |
        Where-Object {$_.membertype -match 'property'} |
        Select-Object -ExpandProperty name

    foreach ($record in $data) {
        & $ScriptBlock $PropertyNames $record
    }
}