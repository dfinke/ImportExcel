# Get-ExcelSheets

param(
    [Parameter(Mandatory)]
    $path
)


$hash = @{ }

$e = Open-ExcelPackage $path

foreach ($sheet in $e.workbook.worksheets) {
    $hash[$sheet.name] = Import-Excel -ExcelPackage $e -WorksheetName $sheet.name
}

Close-ExcelPackage $e -NoSave

$hash