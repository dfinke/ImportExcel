function Read-Excel {
    <#
        .SYNOPSIS
        .EXAMPLE
    #>
    param(
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        $Path,
        # Don't specify a worksheet name and all sheets will be read
        [string[]]$WorksheetName
    )    

    Process {

        if(!$Path) {
            Write-Error "Excel file(s) not specified and are required"
            return
        }

        if (!$WorksheetName) {
            $WorksheetName = Get-ExcelSheetInfo $Path | Select-Object -ExpandProperty Name
        }

        foreach ($sheetname in $WorksheetName) {
            Import-Excel -Path $Path -WorksheetName $sheetname
        }
    }
}
