$result = Read-Excel $PSScriptRoot\yearlySales.xlsx August, October

$result | Group-Object month
"Total rows=" + $result.Count