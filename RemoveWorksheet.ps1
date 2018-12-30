Function Remove-WorkSheet {
    param(
        [Parameter(Mandatory)]
        $Path,
        [Parameter(Mandatory)]
        [String[]]$WorksheetName,
        [Switch]$Show
    )

    $pkg = Open-ExcelPackage -Path $Path

    if ($pkg) {
        foreach ($wsn in $WorksheetName) {
            $pkg.Workbook.Worksheets.Delete($wsn)
        }

        Close-ExcelPackage -ExcelPackage $pkg -Show:$Show
    }
}