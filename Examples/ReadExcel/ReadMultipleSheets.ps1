$result = Read-Excel $PSScriptRoot\yearlySales.xlsx

$result | Group-Object month
"Total rows=" + $result.Count