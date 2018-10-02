﻿Function Remove-WorkSheet {
    Param (
        $Path,
        $WorksheetName
    )

    $Path = (Resolve-Path -Path $path).ProviderPath

    $Excel = New-Object -TypeName OfficeOpenXml.ExcelPackage $Path

    $workSheet = $Excel.Workbook.Worksheets[$WorkSheetName]

    if($workSheet) {
        if($Excel.Workbook.Worksheets.Count -gt 1) {
            $Excel.Workbook.Worksheets.Delete($workSheet)
        } else {
            throw "Cannot delete $WorksheetName. A workbook must contain at least one visible worksheet"
        }

    } else {
        throw "$WorksheetName not found"
    }

    $Excel.Save()
    $Excel.Dispose()
}


Import-Module .\ImportExcel.psd1 -Force

$names = Get-ExcelSheetInfo C:\Temp\testDelete.xlsx
$names | Foreach-Object { Remove-WorkSheet C:\Temp\testDelete.xlsx $_.Name}

##Remove-WorkSheet C:\Temp\testDelete.xlsx sheet6