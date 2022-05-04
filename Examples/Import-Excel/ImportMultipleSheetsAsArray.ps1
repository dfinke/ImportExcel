Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force

$xlfile = "$PSScriptRoot\yearlySales.xlsx"

$result = Import-Excel -Path $xlfile -WorksheetName * -Raw

$result | Measure-Object