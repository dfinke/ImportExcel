function Remove-Worksheet {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        #    [Parameter(ValueFromPipelineByPropertyName)]
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Path')]
        $FullName,
        [String[]]$WorksheetName = "Sheet1",
        [Switch]$Show
    )

    Process {
        if (!$FullName) {
            throw "Remove-Worksheet requires the and Excel file"
        }

        $pkg = Open-ExcelPackage -Path $FullName

        if ($pkg) {
            foreach ($wsn in $WorksheetName) {
                if ($PSCmdlet.ShouldProcess($FullName,"Remove Sheet $wsn")) {
                    $pkg.Workbook.Worksheets.Delete($wsn)
                }
            }
            Close-ExcelPackage -ExcelPackage $pkg -Show:$Show
        }
    }
}