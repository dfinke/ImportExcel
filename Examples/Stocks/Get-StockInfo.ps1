function Get-StockInfo {
    param(
        [Parameter(Mandatory)]
        $symbols,
        [ValidateSet('open', 'close', 'high', 'low', 'avgTotalVolume')]
        $dataPlot = "close"
    )

    $xlfile = "$env:TEMP\stocks.xlsx"
    Remove-Item -Path $xlfile -ErrorAction Ignore

    $result = Invoke-RestMethod "https://api.iextrading.com/1.0/stock/market/batch?symbols=$($symbols)&types=quote&last=1"

    $ecd = New-ExcelChartDefinition -Row 1 -Column 1 -SeriesHeader $dataPlot `
        -XRange symbol -YRange $dataPlot `
        -Title "$($dataPlot)`r`n As Of $((Get-Date).ToShortDateString())"

    $(foreach ($name in $result.psobject.Properties.name) {
            $result.$name.quote
        }) | Export-Excel $xlfile -AutoNameRange -AutoSize -Show -ExcelChartDefinition $ecd -StartRow 21 -StartColumn 2
}
