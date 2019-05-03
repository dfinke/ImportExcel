<#
  .Synopsis

#>
[CmdletBinding()]
param (
    [parameter(ValueFromPipeline = $true)]
    $Path      = $PWD,
    $xlfile    = "$env:TEMP\ScriptAnalyzer.xlsx",
    $ChartType = 'BarClustered' ,
    [switch]$Quiet
)

begin {
    Remove-Item -Path $xlfile -ErrorAction SilentlyContinue
    $xlparams = @{
        Path           = $xlfile
        WorksheetName  = 'FullResults'
        AutoSize       = $true
        AutoFilter     = $true
        Activate       = $true
        Show           = (-not $Quiet)
    }
    $pivotParams = @{
        PivotTableName = 'BreakDown'
        PivotData      = @{RuleName = 'Count' }
        PivotRows      = 'Severity', 'RuleName'
        PivotColumns   = 'Location'
        PivotTotals    = 'Rows'
    }
    $dirsToProcess     = @()
}
process {
    if     ($path.fullName) {$dirsToProcess += $path.fullName}
    elseif ($path.path)     {$dirsToProcess += $path.Path}
    else                    {$dirsToProcess += $path}
}

end {
    $pivotParams['-PivotChartDefinition'] = New-ExcelChartDefinition -ChartType $chartType -Column $dirsToProcess.Count -Title "Script analysis" -LegendBold
    $xlparams['PivotTableDefinition']     = New-PivotTableDefinition @pivotParams

    $dirsToProcess | ForEach-Object {
        $dirName = (Resolve-Path -Path $_) -replace "^.*\\(.*?)\\(.*?)$", '$1-$2'
        Write-Progress -Activity "Running Script Analyzer" -CurrentOperation $dirName
        Invoke-ScriptAnalyzer -Path $_ -ErrorAction SilentlyContinue |
            Add-Member -MemberType NoteProperty -Name Location -Value $dirName -PassThru
    }   |       Export-Excel @xlparams
    Write-Progress -Activity "Running Script Analyzer" -Completed
}
