function Pivot {
    [CmdletBinding()]
    param(
        $targetData,
        $pivotRows,
        $pivotData,
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType="Pie"
    )

    if(!$pivotData) {$pivotData=$pivotRows}

    $xlFile = [System.IO.Path]::GetTempFileName() -replace "tmp","xlsx"
    Write-Verbose $xlFile

    $targetData |
        Export-Excel -Path $xlFile `
            -IncludePivotTable -IncludePivotChart `
            -PivotRows $pivotRows -PivotData $pivotData `
            -ChartType $ChartType `
            -Show
}