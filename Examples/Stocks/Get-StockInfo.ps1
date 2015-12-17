function Get-StockInfo {
    param(
        $stock,
        [datetime]$startDate,
        [datetime]$endDate
    )

    Process {

    if(!$endDate) { $endDate = $startDate}

    $baseUrl = "http://query.yahooapis.com/v1/public/yql?q="
    $q = @"
select * from yahoo.finance.historicaldata where symbol = "$($stock)" and startDate = "$($startDate.ToString('yyyy-MM-dd'))" and endDate = "$($endDate.ToString('yyyy-MM-dd'))"
"@
    $suffix = "&env=store://datatables.org/alltableswithkeys&format=json"
    $r=Invoke-RestMethod ($baseUrl + $q + $suffix)
    $r.query.results.quote

    }
}