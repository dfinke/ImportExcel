function Export-StocksToExcel {
    param(
        [string]$symbols,
        [ValidateSet("Open", "High", "Low", "Close", "Volume")]
        $measure = "Open"
    )

    $xl = Join-Path ([IO.Path]::GetTempPath()) 'Stocks.xlsx'

    Remove-Item $xl -ErrorAction SilentlyContinue

    $r = Invoke-RestMethod "https://azfnstockdata-fn83fffd32.azurewebsites.net/api/GetQuoteChart?symbol=$($symbols)"

    $chartColumn = $symbols.Split(',').count + 2
    $ptd = New-PivotTableDefinition `
        -SourceWorkSheet Sheet1 `
        -PivotTableName result `
        -PivotData @{$measure = 'sum'} `
        -PivotRows date `
        -PivotColumns symbol `
        -ChartType Line `
        -ChartTitle "Stock - $measure " `
        -IncludePivotChart -NoTotalsInPivot -ChartColumn $chartColumn -ChartRow 3 -Activate

    $r | Sort-Object Date, symbol | Export-Excel $xl -PivotTableDefinition $ptd -AutoSize -AutoFilter -Show
}

# Export-StocksToExcel -symbols 'ibm,aapl,msft' -measure High