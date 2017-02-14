Function Get-ExcelSheetInfo {
    <#
    .SYNOPSIS
        Get worksheet names and their indices of an Excel workbook.

    .DESCRIPTION
        The Get-ExcelSheetInfo cmdlet gets worksheet names and their indices of an Excel workbook.

    .PARAMETER Path
        Specifies the path to the Excel file. This parameter is required.

    .EXAMPLE
        Get-ExcelSheetInfo .\Test.xlsx

    .NOTES
        CHANGELOG
        2016/01/07 Added Created by Johan Akerstrom (https://github.com/CosmosKey)

    .LINK
        https://github.com/dfinke/ImportExcel

    #>

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