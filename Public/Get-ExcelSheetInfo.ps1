function Get-ExcelSheetInfo {
    [CmdletBinding()]
    param(
        [Alias('FullName')]
        [Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, Mandatory=$true)]
        $Path
    )
    process {
        $Path = (Resolve-Path $Path).ProviderPath

        $stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Path,'Open','Read','ReadWrite'
        $xl = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $stream
        $workbook  = $xl.Workbook

        if ($workbook -and $workbook.Worksheets) {
            $workbook.Worksheets |
                Select-Object -Property name,index,hidden,@{
                    Label = 'Path'
                    Expression = {$Path}
                }
        }

        $stream.Close()
        $stream.Dispose()
        $xl.Dispose()
        $xl = $null
    }
}