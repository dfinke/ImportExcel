
function New-PivotTableDefinition {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',Justification='Does not change system State')]
    param(
        [Parameter(Mandatory)]
        [Alias("PivtoTableName")]#Previous typo - use alias to avoid breaking scripts
        $PivotTableName,
        $SourceWorksheet,
        $SourceRange,
        $PivotRows,
        [hashtable]$PivotData,
        $PivotColumns,
        $PivotFilter,
        [Switch]$PivotDataToColumn,
        [ValidateSet("Both", "Columns", "Rows", "None")]
        [String]$PivotTotals = "Both",
        [Switch]$NoTotalsInPivot,
        [String]$GroupDateRow,
        [String]$GroupDateColumn,
        [OfficeOpenXml.Table.PivotTable.eDateGroupBy[]]$GroupDatePart,
        [String]$GroupNumericRow,
        [String]$GroupNumericColumn,
        [double]$GroupNumericMin = 0 ,
        [double]$GroupNumericMax = [Double]::MaxValue  ,
        [double]$GroupNumericInterval = 100  ,
        [string]$PivotNumberFormat,
        [OfficeOpenXml.Table.TableStyles]$PivotTableStyle,
        [Parameter(ParameterSetName = 'ChartbyDef', Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        $PivotChartDefinition,
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Switch]$IncludePivotChart,
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [String]$ChartTitle,
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [int]$ChartHeight = 400 ,
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [int]$ChartWidth = 600,
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Int]$ChartRow = 0 ,
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Int]$ChartColumn = 4,
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Int]$ChartRowOffSetPixels = 0 ,
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Int]$ChartColumnOffSetPixels = 0,
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = 'Pie',
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Switch]$NoLegend,
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Switch]$ShowCategory,
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Switch]$ShowPercent,
        [switch]$Activate
    )
    $validDataFuntions = [system.enum]::GetNames([OfficeOpenXml.Table.PivotTable.DataFieldFunctions])

    if ($PivotData.values.Where( {$_ -notin $validDataFuntions}) ) {
        Write-Warning -Message ("Pivot data functions might not be valid, they should be one of " + ($validDataFuntions -join ", ") + ".")
    }

    $parameters = @{} + $PSBoundParameters
    if ($NoTotalsInPivot) {
        $parameters.Remove('NoTotalsInPivot')
        $parameters["PivotTotals"] = "None"
    }
    if  ($PSBoundParameters.ContainsKey('ChartType') -and -not $PSBoundParameters.ContainsKey('IncludePivotChart')) {
        $parameters['IncludePivotChart'] = $true
    }
    $parameters.Remove('PivotTableName')
    if ($PivotChartDefinition) {
        $parameters.PivotChartDefinition.XRange = $null
        $parameters.PivotChartDefinition.YRange = $null
        $parameters.PivotChartDefinition.SeriesHeader = $null
    }
    @{$PivotTableName = $parameters}
}
