Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force

$xlfile = "$PSScriptRoot\yearlySales.xlsx"

$result = Import-Excel $xlfile * -NotAsDictionary

$result | Measure-Object