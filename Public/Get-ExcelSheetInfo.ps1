function Get-ExcelSheetInfo {
    [CmdletBinding()]
    param(
        [Alias('FullName')]
        [Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true, Mandatory = $true)]
        $Path
    )
    process {
        $Path = (Resolve-Path $Path).ProviderPath

        $pkg = Open-ExcelPackage -Path $Path
        $workbook = $pkg.Workbook

        if ($workbook -and $workbook.Worksheets) {
            $workbook.Worksheets | 
                Select-Object -Property Name, Index, Hidden, Dimension, Tables, @{Name = 'Path'; Expression = { $Path } }
        }

        Close-ExcelPackage -ExcelPackage $pkg -NoSave
    }
}