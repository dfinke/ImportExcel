# Read all the xlsx files in the salesData folder
$result = Get-ChildItem $PSScriptRoot\salesData | Read-Excel

$result | Group-Object month
"Total rows=" + $result.Count