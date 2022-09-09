function ConvertTo-Excel {
    param(
        $Path,
        [System.Collections.IDictionary]$targetData
    )

    $column = 1
    foreach ($key in $targetData.Keys) {
        $cityData[$key] | Export-Excel $xlfile -StartColumn ($column++) -HeaderName $key -AutoSize
    }
}

$cityData = [Ordered]@{}

$cityData.City = "New York City", "Paris", "Barcelona", "Rome"
$cityData.Country = "United States", "France", "Spain", "Italy"
$cityData.Population = 8600000, 2141000, 5515000, 2873000

$xlfile = "$PSScriptRoot/test.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

ConvertTo-Excel $xlfile $cityData

. $xlfile