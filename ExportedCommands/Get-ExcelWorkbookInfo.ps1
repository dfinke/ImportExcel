Function Get-ExcelWorkbookInfo {
    [CmdletBinding()]
    Param (
        [Alias('FullName')]
        [Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, Mandatory=$true)]
        [String]$Path
    )

    Process {
        Try {
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
        Catch {
            throw "Failed retrieving Excel workbook information for '$Path': $_"
        }
    }
}
