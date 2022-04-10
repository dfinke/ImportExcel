function Read-Excel {
    <#
        .SYNOPSIS
        Read an Excel file into PowerShell
        .DESCRIPTION
        Supports an option to read a single sheet or a list of sheets, or all the sheets

        .EXAMPLE
        # Read all the sales sheets
        Read-Excel "./yearlySales.xlsx"

        .EXAMPLE
        # Read two sales data sheets april and may
        Read-Excel "./yearlySales.xlsx" april, may

        .EXAMPLE
        # Read all the sheets from all the Excel files in the current directory
        dir *.xlsx | Read-Excel
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
