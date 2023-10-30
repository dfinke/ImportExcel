function Select-Worksheet {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'Package', Position = 0)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        [Parameter(Mandatory = $true, ParameterSetName = 'Workbook')]
        [OfficeOpenXml.ExcelWorkbook]$ExcelWorkbook,
        [Parameter(ParameterSetName='Package')]
        [Parameter(ParameterSetName='Workbook')]
        [string]$WorksheetName,
        [Parameter(ParameterSetName='Sheet',Mandatory=$true)]
        [OfficeOpenXml.ExcelWorksheet]$ExcelWorksheet
    )
    #if we were given a package, use its workbook
    if      ($ExcelPackage   -and -not $ExcelWorkbook) {$ExcelWorkbook  = $ExcelPackage.Workbook}
    #if we now have workbook, get the worksheet; if we were given a sheet get the workbook
    if      ($ExcelWorkbook  -and $WorksheetName)      {$ExcelWorksheet = $ExcelWorkbook.Worksheets[$WorksheetName]}
    elseif  ($ExcelWorksheet -and -not $ExcelWorkbook) {$ExcelWorkbook  = $ExcelWorksheet.Workbook ; }
    #if we didn't get to a worksheet give up. If we did set all works sheets to not selected and then the one we want to selected.
    if (-not $ExcelWorksheet) {Write-Warning -Message "The worksheet $WorksheetName was not found." ; return }
    else {
        foreach ($w in $ExcelWorkbook.Worksheets) {$w.View.TabSelected = $false}
        $ExcelWorksheet.View.TabSelected = $true
    }
}
