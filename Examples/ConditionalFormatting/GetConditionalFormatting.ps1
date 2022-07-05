try { Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 } catch { throw ; return}

$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
$excel = Open-ExcelPackage -Path $xlSourcefile

$excel.Workbook.Worksheets | ForEach-Object {
    $_.ConditionalFormatting | ForEach-Object {
        Write-Host "Add-ConditionalFormatting -Worksheet `$excel[""$worksheetName""]  -Range '$($_.Address)'  -ConditionValue '$($_.Formula)' -RuleType $($_.Type) "
    }
}
