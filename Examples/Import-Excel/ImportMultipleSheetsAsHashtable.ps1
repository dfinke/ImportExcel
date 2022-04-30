Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force

$xlfile = "$PSScriptRoot\yearlySales.xlsx"

$result = Import-Excel $xlfile *

foreach ($sheet in $result.Values) {
    $sheet
}