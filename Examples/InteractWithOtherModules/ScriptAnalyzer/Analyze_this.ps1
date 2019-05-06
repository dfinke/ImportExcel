<#
  .Synopsis
    Runs PsScriptAnalyzer against one or more folders and pivots the results to form a report.

  .Example
    Analyze_this.ps1
    Invokes script analyzer on the current directory; creates a file in $env:temp and opens it in Excel
  .Example
    Analyze_this.ps1 -xlfile ..\mymodule.xlsx -quiet
    Invokes script analyzer on the current directory; creates a file in the parent directory but does not open it
  .Example
    "." , (dir 'C:\Program Files\WindowsPowerShell\Modules\ImportExcel\') | .\examples\ScriptAnalyzer\Analyze_this.ps1
    run from a developemnt directory for importExcel it will produce a report for that directory compared against installed versions
    this creates the file in the default location and opens it
#>
[CmdletBinding()]
param (
    [parameter(ValueFromPipeline = $true)]
    $Path          = $PWD,
    $xlfile        = "$env:TEMP\ScriptAnalyzer.xlsx",
    $ChartType     = 'BarClustered' ,
    $PivotColumns  = 'Location',
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
    $pivotParams['-PivotChartDefinition'] = New-ExcelChartDefinition -ChartType $chartType -Column (1 + $dirsToProcess.Count) -Title "Script analysis" -LegendBold
    $xlparams['PivotTableDefinition']     = New-PivotTableDefinition @pivotParams

    $dirsToProcess | ForEach-Object {
        $dirName = (Resolve-Path -Path $_) -replace "^.*\\(.*?)\\(.*?)$", '$1-$2'
        Write-Progress -Activity "Running Script Analyzer" -CurrentOperation $dirName
        Invoke-ScriptAnalyzer -Path $_ -ErrorAction SilentlyContinue |
            Add-Member -MemberType NoteProperty -Name Location -Value $dirName -PassThru
    }   |       Export-Excel @xlparams
    Write-Progress -Activity "Running Script Analyzer" -Completed
}
