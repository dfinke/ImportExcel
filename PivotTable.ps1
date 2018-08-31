function Add-PivotTable {
    [cmdletbinding(defaultParameterSetName='ChartbyParams')]
    <#
      .Synopsis
        Adds a Pivot table (and optional pivot chart) to a workbook
      .Description
        If the pivot table already exists, the source data will be updated.
    #>
        param (
            #Name for the new Pivot table - this will be the name of a sheet in the workbook
            [Parameter(Mandatory = $true)]
            [string]$PivotTableName,
            #By default a pivot table will be created on its own sheet, but it can be created on an existing sheet by giving the address where the top left corner of the table should go. (Allow two rows for the filter if one is used.)
            [OfficeOpenXml.ExcelAddressBase]
            $Address,
            #An excel package object for the workbook.
            $ExcelPackage,
            #Worksheet where the data is found
            $SourceWorkSheet,
            #Address range in the worksheet e.g "A10:F20" - the first row must be column names: if not specified the whole sheet will be used.
            $SourceRange,
            #Fields to set as rows in the Pivot table
            $PivotRows,
            #A hash table in form "FieldName"="Function", where function is one of
            #Average, Count, CountNums, Max, Min, Product, None, StdDev, StdDevP, Sum, Var, VarP
            $PivotData,
            #Fields to set as columns in the Pivot table
            $PivotColumns,
            #Fields to use to filter in the Pivot table
            $PivotFilter,
            [Switch]$PivotDataToColumn,
            #By default Pivot tables have Totals for each Row (on the right) and for each column at the bottom. This allows just one or neither to be selected.
            [ValidateSet("Both","Columns","Rows","None")]
            [String]$PivotTotals = "Both",
            #Included for compatibility - equivalent to -PivotTotals "None"
            [Switch]$NoTotalsInPivot,
            #Number format to apply to the data cells in the Pivot table
            [string]$PivotNumberFormat,
            #Apply a table style to the PivotTable
            [OfficeOpenXml.Table.TableStyles]$PivotTableSyle,
            #Use a chart definition instead of specifying chart settings one by one
            [Parameter(ParameterSetName='ChartbyDef', Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
            $PivotChartDefinition,
            #If specified a chart Will be included.
            [Parameter(ParameterSetName='ChartbyParams')]
            [Switch]$IncludePivotChart,
            #Optional title for the pivot chart, by default the title omitted.
            [Parameter(ParameterSetName='ChartbyParams')]
            [String]$ChartTitle = "",
            #Height of the chart in Pixels (400 by default)
            [Parameter(ParameterSetName='ChartbyParams')]
            [int]$ChartHeight = 400 ,
            #Width of the chart in Pixels (600 by default)
            [Parameter(ParameterSetName='ChartbyParams')]
            [int]$ChartWidth = 600,
            #Cell position of the top left corner of the chart, there will be this number of rows above the top edge of the chart (default is 0, chart starts at top edge of row 1).
            [Parameter(ParameterSetName='ChartbyParams')]
            [Int]$ChartRow = 0 ,
            #Cell position of the top left corner of the chart, there will be this number of cells to the left of the chart (default is 4, chart starts at left edge of column E)
            [Parameter(ParameterSetName='ChartbyParams')]
            [Int]$ChartColumn = 4,
            #Vertical offset of the chart from the cell corner.
            [Parameter(ParameterSetName='ChartbyParams')]
            [Int]$ChartRowOffSetPixels = 0 ,
            [Parameter(ParameterSetName='ChartbyParams')]
            #Horizontal offset of the chart from the cell corner.
            [Int]$ChartColumnOffSetPixels = 0,
            #Type of chart
            [Parameter(ParameterSetName='ChartbyParams')]
            [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = 'Pie',
            #If specified hides the chart legend
            [Parameter(ParameterSetName='ChartbyParams')]
            [Switch]$NoLegend,
            #if specified attaches the category to slices in a pie chart : not supported on all chart types, this may give errors if applied to an unsupported type.
            [Parameter(ParameterSetName='ChartbyParams')]
            [Switch]$ShowCategory,
            #If specified attaches percentages to slices in a pie chart.
            [Parameter(ParameterSetName='ChartbyParams')]
            [Switch]$ShowPercent,
            #If there is already content in the workbook the sheet with the Pivot table will not be active UNLESS Activate is specified
            [switch]$Activate,
            #Return the pivot table so it can be customized
            [Switch]$PassThru
        )
        if ($PivotTableName.length -gt 250) {
            Write-warning -Message "Pivot table name will be truncated"
            $PivotTableName = $PivotTableName.Substring(0,250)
        }
        if ($Address) {
            [OfficeOpenXml.ExcelWorksheet]$wsPivot = $address.Worksheet
        }
        else {
            [OfficeOpenXml.ExcelWorksheet]$wsPivot = Add-WorkSheet -ExcelPackage $ExcelPackage -WorksheetName $pivotTableName -Activate:$Activate
            if ($wsPivot.Name -ne $PivotTableName) {Write-Warning -Message "The Worksheet name for the pivot table does not match the table name '$PivotTableName'; probably because excess or illegal characters were removed." }
            if ($PivotFilter) {$Address =  $wsPivot.Cells["A3"]} else { $Address =  $wsPivot.Cells["A1"]}
        }
        #if the pivot doesn't exist, create it.
        if (-not $wsPivot.PivotTables[$pivotTableName] ) {
            try {
                #Accept a string or a worksheet object as $Source Worksheet.
                if     ($SourceWorkSheet -is [string]) {
                        $SourceWorkSheet = $ExcelPackage.Workbook.Worksheets.where( {$_.name -Like $SourceWorkSheet})[0]
                }
                elseif ($SourceWorkSheet -is [int])    {
                        $SourceWorkSheet = $ExcelPackage.Workbook.Worksheets[$SourceWorkSheet]
                }
                if     ($SourceWorkSheet -isnot  [OfficeOpenXml.ExcelWorksheet]) {Write-Warning -Message "Could not find source Worksheet for pivot-table '$pivotTableName'." ; return }
                else {

                    if (-not $SourceRange) { $SourceRange = $SourceWorkSheet.Dimension.Address}
                    $pivotTable = $wsPivot.PivotTables.Add($Address, $SourceWorkSheet.Cells[ $SourceRange], $pivotTableName)
                }
                foreach ($Row in $PivotRows) {
                    try {$null = $pivotTable.RowFields.Add($pivotTable.Fields[$Row]) }
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
                if     ($NoTotalsInPivot) {$PivotTotals = "None" }
                if     ($PivotTotals -eq "None" -or $PivotTotals -eq "Columns") { $pivotTable.RowGrandTotals   = $false }
                elseif ($PivotTotals -eq "Both" -or $PivotTotals -eq "Rows")    { $pivotTable.RowGrandTotals   = $true  }
                if     ($PivotTotals -eq "None" -or $PivotTotals -eq "Rows")    { $pivotTable.ColumGrandTotals = $false }   # Epplus spelling mistake, not mine!
                elseif ($PivotTotals -eq "Both" -or $PivotTotals -eq "Columns") { $pivotTable.ColumGrandTotals = $true  }
                if     ($PivotDataToColumn ) { $pivotTable.DataOnRows = $false }
                if     ($PivotTableSyle)     { $pivotTable.TableStyle = $PivotTableSyle}
            }
            catch {Write-Warning -Message "Failed adding PivotTable '$pivotTableName': $_"}
        }
        else {
            Write-Warning -Message "Pivot table defined in $($pivotTableName) already exists, only the data range will be changed."
            $pivotTable = $wsPivot.PivotTables[$pivotTableName]
            $pivotTable.CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.ref = $SourceRange
        }

        #Create the chart if it doesn't exist, leave alone if it does.
        if ($IncludePivotChart -and -not $wsPivot.Drawings["Chart$pivotTableName"] ) {
            try   {Add-ExcelChart -PivotTable $pivotTable -ChartType $ChartType -Width $ChartWidth -Height $ChartHeight -Row $ChartRow -Column $ChartColumn -RowOffSetPixels $ChartRowOffSetPixels -ColumnOffSetPixels $ChartColumnOffSetPixels -Title $ChartTitle -NoLegend:$NoLegend -ShowCategory:$ShowCategory -ShowPercent:$ShowPercent  }
            catch {Write-Warning -Message "Failed adding chart for pivotable '$pivotTableName': $_"}
        }
        elseif ($PivotChartDefinition -and -not $wsPivot.Drawings["Chart$pivotTableName"]) {
            $params = @{PivotTable= $pivotTable }
            $PivotChartDefinition.PSObject.Properties | ForEach-Object {if ( $null -ne $_.value) {$params[$_.name] = $_.value}}
            Add-ExcelChart @params
        }
        if ($PassThru) {return $pivotTable}
    }

function New-PivotTableDefinition {
        <#
          .Synopsis
            Creates Pivot table definitons for Export-Excel
          .Description
            Export-Excel allows a single Pivot table to be defined using the parameters -IncludePivotTable, -PivotColumns -PivotRows,
            =PivotData, -PivotFilter, -PivotTotals, -PivotDataToColumn, -IncludePivotChart and -ChartType.
            Its -PivotTableDefintion paramater allows multiple pivot tables to be defined, with additional parameters.
            New-PivotTableDefinition is a convenient way to build these definitions.
          .Example
            $pt  = New-PivotTableDefinition -PivotTableName "PT1" -SourceWorkSheet "Sheet1" -PivotRows "Status"  -PivotData @{Status='Count' } -PivotFilter 'StartType' -IncludePivotChart  -ChartType BarClustered3D
            $Pt += New-PivotTableDefinition -PivotTableName "PT2" -SourceWorkSheet "Sheet2" -PivotRows "Company" -PivotData @{Company='Count'} -IncludePivotChart  -ChartType PieExploded3D  -ShowPercent -ChartTitle "Breakdown of processes by company"
            Get-Service | Select-Object    -Property Status,Name,DisplayName,StartType | Export-Excel -Path .\test.xlsx -AutoSize
            Get-Process | Select-Object    -Property Name,Company,Handles,CPU,VM       | Export-Excel -Path .\test.xlsx -AutoSize -WorksheetName 'sheet2'
            $excel = Export-Excel -Path .\test.xlsx -PivotTableDefinition $pt -Show

            This is a re-work of one of the examples in Export-Excel - instead of writing out the pivot definition hash table it is built by calling New-PivotTableDefinition.
        #>
        param(
            [Parameter(Mandatory)]
            [Alias("PivtoTableName")]#Previous typo - use alias to avoid breaking scripts
            $PivotTableName,
            #Worksheet where the data is found
            $SourceWorkSheet,
            #Address range in the worksheet e.g "A10:F20" - the first row must be column names: if not specified the whole sheet will be used/
            $SourceRange,
            #Fields to set as rows in the Pivot table
            $PivotRows,
            #A hash table in form "FieldName"="Function", where function is one of
            #Average, Count, CountNums, Max, Min, Product, None, StdDev, StdDevP, Sum, Var, VarP
            [hashtable]$PivotData,
            #Fields to set as columns in the Pivot table
            $PivotColumns,
            #Fields to use to filter in the Pivot table
            $PivotFilter,
            [Switch]$PivotDataToColumn,
            #By default Pivot tables have Totals for each Row (on the right) and for each column at the bottom. This allows just one or neither to be selected.
            [ValidateSet("Both","Columns","Rows","None")]
            [String]$PivotTotals = "Both",
            #Included for compatibility - equivalent to -PivotTotals "None"
            [Switch]$NoTotalsInPivot,
            #Number format to apply to the data cells in the Pivot table
            [string]$PivotNumberFormat,
            #Apply a table style to the PivotTable
            [OfficeOpenXml.Table.TableStyles]$PivotTableSyle,
            #Use a chart definition instead of specifying chart settings one by one
            [Parameter(ParameterSetName='ChartbyDef', Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
            $PivotChartDefinition,
            #If specified a chart Will be included.
            [Parameter(ParameterSetName='ChartbyParams')]
            [Switch]$IncludePivotChart,
            #Optional title for the pivot chart, by default the title omitted.
            [Parameter(ParameterSetName='ChartbyParams')]
            [String]$ChartTitle,
            #Height of the chart in Pixels (400 by default)
            [Parameter(ParameterSetName='ChartbyParams')]
            [int]$ChartHeight = 400 ,
            #Width of the chart in Pixels (600 by default)
            [Parameter(ParameterSetName='ChartbyParams')]
            [int]$ChartWidth = 600,
            #Cell position of the top left corner of the chart, there will be this number of rows above the top edge of the chart (default is 0, chart starts at top edge of row 1).
            [Parameter(ParameterSetName='ChartbyParams')]
            [Int]$ChartRow = 0 ,
            #Cell position of the top left corner of the chart, there will be this number of cells to the left of the chart (default is 4, chart starts at left edge of column E)
            [Parameter(ParameterSetName='ChartbyParams')]
            [Int]$ChartColumn = 4,
            #Vertical offset of the chart from the cell corner.
            [Parameter(ParameterSetName='ChartbyParams')]
            [Int]$ChartRowOffSetPixels = 0 ,
            #Horizontal offset of the chart from the cell corner.
            [Parameter(ParameterSetName='ChartbyParams')]
            [Int]$ChartColumnOffSetPixels = 0,
            #Type of chart
            [Parameter(ParameterSetName='ChartbyParams')]
            [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = 'Pie',
            #If specified hides the chart legend
            [Parameter(ParameterSetName='ChartbyParams')]
            [Switch]$NoLegend,
            #if specified attaches the category to slices in a pie chart : not supported on all chart types, this may give errors if applied to an unsupported type.
            [Parameter(ParameterSetName='ChartbyParams')]
            [Switch]$ShowCategory,
            #If specified attaches percentages to slices in a pie chart.
            [Parameter(ParameterSetName='ChartbyParams')]
            [Switch]$ShowPercent,
            #If there is already content in the workbook the sheet with the Pivot table will not be active UNLESS Activate is specified
            [switch]$Activate
        )
        $validDataFuntions = [system.enum]::GetNames([OfficeOpenXml.Table.PivotTable.DataFieldFunctions])

        if ($PivotData.values.Where({$_ -notin $validDataFuntions}) ) {
            Write-Warning -Message ("Pivot data functions might not be valid, they should be one of " + ($validDataFuntions -join ", ") + ".")
        }

        $parameters = @{} + $PSBoundParameters
        if ($NoTotalsInPivot) {
            $parameters.Remove('NoTotalsInPivot')
            $parameters["PivotTotals"] = "None"
        }
        $parameters.Remove('PivotTableName')
        if ($PivotChartDefinition) {
            $parameters.PivotChartDefinition.XRange       = $null
            $parameters.PivotChartDefinition.YRange       = $null
            $parameters.PivotChartDefinition.SeriesHeader = $null
        }
        @{$PivotTableName = $parameters}
    }
