function Add-PivotTable {
    [CmdletBinding(defaultParameterSetName = 'ChartbyParams')]
    [OutputType([OfficeOpenXml.Table.PivotTable.ExcelPivotTable])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$PivotTableName,
        [OfficeOpenXml.ExcelAddressBase]
        $Address,
        $ExcelPackage,
        $SourceWorksheet,
        $SourceRange,
        $PivotRows,
        $PivotData,
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
        [String]$ChartTitle = "",
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
        [switch]$Activate,
        [Switch]$PassThru
    )
    if ($PivotTableName.length -gt 250) {
        Write-warning -Message "PivotTable name will be truncated"
        $PivotTableName = $PivotTableName.Substring(0, 250)
    }
    if ($Address) {
        [OfficeOpenXml.ExcelWorksheet]$wsPivot = $address.Worksheet
    }
    else {
        try {
            if (-not $ExcelPackage) {Write-Warning -message "This combination of Parameters needs to include the ExcelPackage." ; return }
            [OfficeOpenXml.ExcelWorksheet]$wsPivot = Add-Worksheet -ExcelPackage $ExcelPackage -WorksheetName $pivotTableName -Activate:$Activate
            if ($wsPivot.Name -ne $PivotTableName) {Write-Warning -Message "The Worksheet name for the PivotTable does not match the table name '$PivotTableName'; probably because excess or illegal characters were removed." }
            if ($PivotFilter) {$Address = $wsPivot.Cells["A3"]} else { $Address = $wsPivot.Cells["A1"]}
        }
        catch {throw "Could not create the sheet for the PivotTable. $_" }
    }
    #if the pivot doesn't exist, create it.
    if (-not $wsPivot) {throw "There was a problem getting the worksheet for the PivotTable"}
    if (-not $wsPivot.PivotTables[$pivotTableName] ) {
        try {
            #Accept a string or a worksheet object as $SourceWorksheet - we don't need a worksheet if we have a Rangebase .
            if ( $SourceWorksheet -is [string]) {
                $SourceWorksheet = $ExcelPackage.Workbook.Worksheets.where( {$_.name -Like $SourceWorksheet})[0]
            }
            elseif ( $SourceWorksheet -is [int]   ) {
                $SourceWorksheet = $ExcelPackage.Workbook.Worksheets[$SourceWorksheet]
            }
            if (     $SourceRange -is [OfficeOpenXml.Table.ExcelTable]) {$SourceRange = $SourceRange.Address }
            if (     $sourceRange -is [OfficeOpenXml.ExcelRange] -or
                $SourceRange -is [OfficeOpenXml.ExcelAddress]) {
                $pivotTable = $wsPivot.PivotTables.Add($Address, $SourceRange, $pivotTableName)
            }
            elseif (-not $SourceRange) {
                $pivotTable = $wsPivot.PivotTables.Add($Address, $SourceWorksheet.cells[$SourceWorksheet.Dimension.Address], $pivotTableName)
            }
            elseif ($SourceWorksheet -isnot [OfficeOpenXml.ExcelWorksheet]  ) {
                Write-Warning -Message "Could not find source Worksheet for pivot-table '$pivotTableName'." ; return
            }
            elseif (     $SourceRange -is [String] -or $SourceRange -is [OfficeOpenXml.ExcelAddress]) {
                $pivotTable = $wsPivot.PivotTables.Add($Address, $SourceWorksheet.Cells[$SourceRange], $pivotTableName)
            }
            else {Write-warning "Could not create a PivotTable with the Source Range provided."; return}
            foreach ($row in $PivotRows) {
                try {$null = $pivotTable.RowFields.Add($pivotTable.Fields[$row]) }
                catch {Write-Warning -message "Could not add '$row' to Rows in PivotTable $pivotTableName." }
            }
            foreach ($Column in $PivotColumns) {
                try {$null = $pivotTable.ColumnFields.Add($pivotTable.Fields[$Column])}
                catch {Write-Warning -message "Could not add '$Column' to Columns in PivotTable $pivotTableName." }
            }
            if ($PivotData -is [HashTable] -or $PivotData -is [System.Collections.Specialized.OrderedDictionary]) {
                $PivotData.Keys | ForEach-Object {
                    try {
                        $df = $pivotTable.DataFields.Add($pivotTable.Fields[$_])
                        $df.Function = $PivotData.$_
                        if ($PivotNumberFormat) {$df.Format = (Expand-NumberFormat -NumberFormat $PivotNumberFormat)}
                    }
                    catch {Write-Warning -message "Problem adding data fields to PivotTable $pivotTableName." }
                }
            }
            else {
                foreach ($field in $PivotData) {
                    try {
                        $df = $pivotTable.DataFields.Add($pivotTable.Fields[$field])
                        $df.Function = 'Count'
                    }
                    catch {Write-Warning -message "Problem adding data field '$field' to PivotTable $pivotTableName." }
                }
            }
            foreach ( $pFilter in $PivotFilter) {
                try { $null = $pivotTable.PageFields.Add($pivotTable.Fields[$pFilter])}
                catch {Write-Warning -message "Could not add '$pFilter' to Filter/Page fields in PivotTable $pivotTableName." }
            }
            if ($NoTotalsInPivot) {$PivotTotals = "None" }
            if ($PivotTotals -eq "None" -or $PivotTotals -eq "Columns") { $pivotTable.RowGrandTotals = $false }
            elseif ($PivotTotals -eq "Both" -or $PivotTotals -eq "Rows") { $pivotTable.RowGrandTotals = $true  }
            if ($PivotTotals -eq "None" -or $PivotTotals -eq "Rows") { $pivotTable.ColumGrandTotals = $false }   # Epplus spelling mistake, not mine!
            elseif ($PivotTotals -eq "Both" -or $PivotTotals -eq "Columns") { $pivotTable.ColumGrandTotals = $true  }
            if ($PivotDataToColumn ) { $pivotTable.DataOnRows = $false }
            if ($PivotTableStyle) { $pivotTable.TableStyle = $PivotTableStyle}
            if ($GroupNumericRow) {
                $r = $pivotTable.RowFields.Where( {$_.name -eq $GroupNumericRow })
                if (-not $r ) {Write-Warning -Message "Could not find a Row field named '$GroupNumericRow'; no numeric grouping will be done."}
                else {$r.AddNumericGrouping($GroupNumericMin, $GroupNumericMax, $GroupNumericInterval)}
            }
            elseif ($GroupNumericColumn) {
                $c = $pivotTable.ColumnFields.Where( {$_.name -eq $GroupNumericColumn })
                if (-not $c ) {Write-Warning -Message "Could not find a Column field named '$GroupNumericColumn'; no numeric grouping will be done."}
                else {$c.AddNumericGrouping($GroupNumericMin, $GroupNumericMax, $GroupNumericInterval)}
            }
            if ($GroupDateRow -and $PSBoundParameters.ContainsKey("GroupDatePart")) {
                $r = $pivotTable.RowFields.Where( {$_.name -eq $GroupDateRow })
                if (-not $r ) {Write-Warning -Message "Could not find a Row field named '$GroupDateRow'; no date grouping will be done."}
                else {$r.AddDateGrouping($GroupDatePart)}
            }
            elseif ($GroupDateColumn -and $PSBoundParameters.ContainsKey("GroupDatePart")) {
                $c = $pivotTable.ColumnFields.Where( {$_.name -eq $GroupDateColumn })
                if (-not $c ) {Write-Warning -Message "Could not find a Column field named '$GroupDateColumn'; no date grouping will be done."}
                else {$c.AddDateGrouping($GroupDatePart)}
            }
        }
        catch {Write-Warning -Message "Failed adding PivotTable '$pivotTableName': $_"}
    }
    else {
        Write-Warning -Message "PivotTable defined in $($pivotTableName) already exists, only the data range will be changed."
        $pivotTable = $wsPivot.PivotTables[$pivotTableName]
        if (-not $SourceRange) { $SourceRange = $SourceWorksheet.Dimension.Address}
        $pivotTable.CacheDefinition.SourceRange =  $SourceWorksheet.cells[$SourceRange]
        #change for epPlus 4.5 -  Previously needed to hack the xml
      # $pivotTable.CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.ref = $SourceRange

     }

    #Create the chart if it doesn't exist, leave alone if it does.
    if ($IncludePivotChart -and -not $wsPivot.Drawings["Chart$pivotTableName"] ) {
        try {Add-ExcelChart -PivotTable $pivotTable -ChartType $ChartType -Width $ChartWidth -Height $ChartHeight -Row $ChartRow -Column $ChartColumn -RowOffSetPixels $ChartRowOffSetPixels -ColumnOffSetPixels $ChartColumnOffSetPixels -Title $ChartTitle -NoLegend:$NoLegend -ShowCategory:$ShowCategory -ShowPercent:$ShowPercent  }
        catch {Write-Warning -Message "Failed adding chart for pivotable '$pivotTableName': $_"}
    }
    elseif ($PivotChartDefinition -and -not $wsPivot.Drawings["Chart$pivotTableName"]) {
        if ($PivotChartDefinition -is [System.Management.Automation.PSCustomObject]) {
            $params = @{PivotTable = $pivotTable }
            $PivotChartDefinition.PSObject.Properties | ForEach-Object {if ( $null -ne $_.value) {$params[$_.name] = $_.value}}
            Add-ExcelChart @params
        }
        elseif ($PivotChartDefinition -is [hashtable] -or $PivotChartDefinition -is [System.Collections.Specialized.OrderedDictionary]) {
            Add-ExcelChart -PivotTable $pivotTable  @PivotChartDefinition
        }
    }
    if ($PassThru) {return $pivotTable}
}
