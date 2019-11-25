function Get-ExcelWorkbookInfo {
    [CmdletBinding()]
    param (
        [Alias('FullName')]
        [Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, Mandatory=$true)]
        [String]$Path
    )

    process {
        try {
            $Path = (Resolve-Path $Path).ProviderPath

            $stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Path,'Open','Read','ReadWrite'
            $xl = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $stream
            $workbook  = $xl.Workbook
            $workbook.Properties

            $stream.Close()
            $stream.Dispose()
            $xl.Dispose()
            $xl = $null
        }
        catch {
            throw "Failed retrieving Excel workbook information for '$Path': $_"
        }
    }
}
