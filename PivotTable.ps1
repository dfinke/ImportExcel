function Add-PivotTable {
    <#
      .Synopsis
        Adds a PivotTable (and optional PivotChart) to a workbook.
      .Description
        If the PivotTable already exists, the source data will be updated.
      .Example
        >
        PS> $excel = Get-Service | Export-Excel -Path test.xlsx -WorksheetName Services -PassThru -AutoSize -DisplayPropertySet -TableName ServiceTable -Title "Services on $Env:COMPUTERNAME"
        Add-PivotTable -ExcelPackage $excel  -PivotTableName ServiceSummary   -SourceRange $excel.Workbook.Worksheets[1].Tables[0].Address -PivotRows Status -PivotData Name -NoTotalsInPivot -Activate
        Close-ExcelPackage $excel -Show

        This exports data to new workbook and creates a table with the data in.
        The Pivot table is added on its own page, the table created in the first command is used as the source for the PivotTable;  which counts the service names in for each Status. At the end the Pivot page is made active.
      .Example
        >
        PS> $chartdef = New-ExcelChartDefinition -Title "Gross and net by city and product"  -ChartType ColumnClustered `
               -Column 11 -Width 500 -Height 360 -YMajorUnit 500 -YMinorUnit 100 -YAxisNumberformat "$#,##0" -LegendPosition Bottom

        $excel = ConvertFrom-Csv    @"
        Product, City, Gross, Net
        Apple, London , 300, 250
        Orange, London , 400, 350
        Banana, London , 300, 200
        Orange, Paris,   600, 500
        Banana, Paris,   300, 200
        Apple, New York, 1200,700
        "@  | Export-Excel  -Path "test.xlsx"  -TableStyle Medium13 -tablename "RawData" -PassThru

        Add-PivotTable -PivotTableName Sales -Address $excel.Workbook.Worksheets[1].Cells["F1"] `
                    -SourceWorkSheet $excel.Workbook.Worksheets[1]  -PivotRows City -PivotColumns Product -PivotData @{Gross="Sum";Net="Sum"} `
                    -PivotNumberFormat "$#,##0.00"  -PivotTotals Both  -PivotTableStyle Medium12 -PivotChartDefinition $chartdef
        Close-ExcelPackage -show $excel


        This script starts by defining a chart. Then it exports some data to an XLSX file and keeps the file open.
        The next step is to add the pivot table, normally this would be on its own sheet in the workbook,
        but here -Address is specified to place it beside the data. The Add-Pivot table is given the chart definition and told to create a tale
        using the City field to create rows, the Product field to create columns and the data should be the sum of the gross field and the sum of the net field;
        grand totals for both gross and net are included for rows (Cities) and columns (product) and the data is explicitly formatted as a currency.
        Note that in the chart definition the number format for the axis does not include any fraction part.
      .Example
        >
        $excel = Convertfrom-csv @"
        Location,OrderDate,quantity
        Boston,1/1/2017,100
        New York,1/21/2017,200
        Boston,1/11/2017,300
        New York,1/9/2017,400
        Boston,1/18/2017,500
        Boston,2/1/2017,600
        New York,2/21/2017,700
        New York,2/11/2017,800
        Boston,2/9/2017,900
        Boston,2/18/2017,1000
        New York,1/1/2018,100
        Boston,1/21/2018,200
        New York,1/11/2018,300
        Boston,1/9/2018,400
        New York,1/18/2018,500
        Boston,2/1/2018,600
        Boston,2/21/2018,700
        New York,2/11/2018,800
        New York,2/9/2018,900
        Boston,2/18/2018,1000
        "@ | Select-Object -Property @{n="OrderDate";e={[datetime]::ParseExact($_.OrderDate,"M/d/yyyy",(Get-Culture))}},
                Location, Quantity |  Export-Excel "test2.xlsx" -PassThru -AutoSize

        Set-ExcelColumn -Worksheet $excel.sheet1 -Column 1 -NumberFormat 'Short Date'

        $pt = Add-PivotTable -PassThru -PivotTableName "ByDate" -Address $excel.Sheet1.cells["F1"] -SourceWorkSheet $excel.Sheet1  -PivotRows location,orderdate -PivotData @{'quantity'='sum'}  -GroupDateRow orderdate -GroupDatePart 'Months,Years' -PivotTotals None
        $pt.RowFields[0].SubtotalTop=$false
        $pt.RowFields[0].Compact=$false
        Close-ExcelPackage $excel -Show

        Here the data contains dates formatted as strings using US format. These
        are converted to DateTime objects before being exported to Excel; the
        "OrderDate" column is formatted with the local short-date style. Then
        the PivotTable is added; it groups information by date and location, the
        date is split into years and then months. No grand totals are displayed.
        The Pivot table object is caught in a variable, and the "Location"
        column has its subtotal moved from the top to the bottom of each location
        section, and the "Compact" option is disabled to prevent "Year" moving
        into the same column as location.
        Finally the workbook is saved and shown in Excel.
    #>
    [cmdletbinding(defaultParameterSetName = 'ChartbyParams')]
    [OutputType([OfficeOpenXml.Table.PivotTable.ExcelPivotTable])]
    param (
        #Name for the new PivotTable - this will be the name of a sheet in the Workbook.
        [Parameter(Mandatory = $true)]
        [string]$PivotTableName,
        #By default, a PivotTable will be created on its own sheet, but it can be created on an existing sheet by giving the address where the top left corner of the table should go. (Allow two rows for the filter if one is used.)
        [OfficeOpenXml.ExcelAddressBase]
        $Address,
        #An excel package object for the workbook.
        $ExcelPackage,
        #Worksheet where the data is found.
        $SourceWorkSheet,
        #Address range in the worksheet e.g "A10:F20" - the first row must be column names: if not specified the whole sheet will be used.
        $SourceRange,
        #Fields to set as rows in the PivotTable.
        $PivotRows,
        #A hash table in form "FieldName"="Function", where function is one of
        #Average, Count, CountNums, Max, Min, Product, None, StdDev, StdDevP, Sum, Var, VarP.
        $PivotData,
        #Fields to set as columns in the PivotTable.
        $PivotColumns,
        #Fields to use to filter in the PivotTable.
        $PivotFilter,
        #If there are multiple data items in a PivotTable, by default they are shown on separate rows; this switch makes them separate columns.
        [Switch]$PivotDataToColumn,
        #Define whether totals should be added to rows, columns neither, or both (the default is both).
        [ValidateSet("Both", "Columns", "Rows", "None")]
        [String]$PivotTotals = "Both",
        #Included for compatibility - equivalent to -PivotTotals "None".
        [Switch]$NoTotalsInPivot,
        #The name of a row field which should be grouped by parts of the date/time (ignored if GroupDateRow is not specified)
        [String]$GroupDateRow,
        #The Part(s) of the date to use in the grouping (ignored if GroupDateRow is not specified)
        [OfficeOpenXml.Table.PivotTable.eDateGroupBy[]]$GroupDatePart,
        #The name of a row field which should be grouped by Number (e.g 0-99, 100-199, 200-299 )
        [String]$GroupNumericRow,
        #The starting point for grouping
        [double]$GroupNumericMin = 0 ,
        #The endpoint for grouping
        [double]$GroupNumericMax = [Double]::MaxValue  ,
        #The interval for grouping
        [double]$GroupNumericInterval = 100  ,
        #Number format to apply to the data cells in the PivotTable.
        [string]$PivotNumberFormat,
        #Apply a table style to the PivotTable.
        [OfficeOpenXml.Table.TableStyles]$PivotTableStyle,
        #Use a chart definition instead of specifying chart settings one by one.
        [Parameter(ParameterSetName = 'ChartbyDef', Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        $PivotChartDefinition,
        #If specified, a chart will be included.
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Switch]$IncludePivotChart,
        #Optional title for the pivot chart, by default the title omitted.
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [String]$ChartTitle = "",
        #Height of the chart in Pixels (400 by default).
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [int]$ChartHeight = 400 ,
        #Width of the chart in Pixels (600 by default).
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [int]$ChartWidth = 600,
        #Cell position of the top left corner of the chart, there will be this number of rows above the top edge of the chart (default is 0, chart starts at top edge of row 1).
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Int]$ChartRow = 0 ,
        #Cell position of the top left corner of the chart, there will be this number of cells to the left of the chart (default is 4, chart starts at left edge of column E).
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Int]$ChartColumn = 4,
        #Vertical offset of the chart from the cell corner.
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Int]$ChartRowOffSetPixels = 0 ,
        [Parameter(ParameterSetName = 'ChartbyParams')]
        #Horizontal offset of the chart from the cell corner.
        [Int]$ChartColumnOffSetPixels = 0,
        #Type of chart; defaults to "Pie".
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = 'Pie',
        #If specified hides the chart legend.
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Switch]$NoLegend,
        #If specified attaches the category to slices in a pie chart : not supported on all chart types, this may give errors if applied to an unsupported type.
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Switch]$ShowCategory,
        #If specified attaches percentages to slices in a pie chart.
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Switch]$ShowPercent,
        #If there is already content in the workbook the sheet with the PivotTable will not be active UNLESS Activate is specified.
        [switch]$Activate,
        #Return the PivotTable so it can be customized.
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
            [OfficeOpenXml.ExcelWorksheet]$wsPivot = Add-WorkSheet -ExcelPackage $ExcelPackage -WorksheetName $pivotTableName -Activate:$Activate
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
            if ( $SourceWorkSheet -is [string]) {
                $SourceWorkSheet = $ExcelPackage.Workbook.Worksheets.where( {$_.name -Like $SourceWorkSheet})[0]
            }
            elseif ( $SourceWorkSheet -is [int]   ) {
                $SourceWorkSheet = $ExcelPackage.Workbook.Worksheets[$SourceWorkSheet]
            }
            if (     $SourceRange -is [OfficeOpenXml.Table.ExcelTable]) {$SourceRange = $SourceRange.Address }
            if (     $sourceRange -is [OfficeOpenXml.ExcelRange] -or
                $SourceRange -is [OfficeOpenXml.ExcelAddress]) {
                $pivotTable = $wsPivot.PivotTables.Add($Address, $SourceRange, $pivotTableName)
            }
            elseif (-not $SourceRange) {
                $pivotTable = $wsPivot.PivotTables.Add($Address, $SourceWorkSheet.cells[$SourceWorkSheet.Dimension.Address], $pivotTableName)
            }
            elseif ($SourceWorkSheet -isnot [OfficeOpenXml.ExcelWorksheet]  ) {
                Write-Warning -Message "Could not find source Worksheet for pivot-table '$pivotTableName'." ; return
            }
            elseif (     $SourceRange -is [String] -or $SourceRange -is [OfficeOpenXml.ExcelAddress]) {
                $pivotTable = $wsPivot.PivotTables.Add($Address, $SourceWorkSheet.Cells[$SourceRange], $pivotTableName)
            }
            else {Write-warning "Could not create a PivotTable with the Source Range provided."; return}
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
            if ($GroupDateRow -and $PSBoundParameters.ContainsKey("GroupDatePart")) {
                $r = $pivotTable.RowFields.Where( {$_.name -eq $GroupDateRow })
                if (-not $r ) {Write-Warning -Message "Could not find a Row field named '$GroupDateRow'; no date grouping will be done."}
                else {$r.AddDateGrouping($GroupDatePart)}
            }
        }
        catch {Write-Warning -Message "Failed adding PivotTable '$pivotTableName': $_"}
    }
    else {
        Write-Warning -Message "PivotTable defined in $($pivotTableName) already exists, only the data range will be changed."
        $pivotTable = $wsPivot.PivotTables[$pivotTableName]
        if (-not $SourceRange) { $SourceRange = $SourceWorkSheet.Dimension.Address}
        $pivotTable.CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.ref = $SourceRange
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

function New-PivotTableDefinition {
    <#
      .Synopsis
        Creates PivotTable definitons for Export-Excel
      .Description
        Export-Excel allows a single PivotTable to be defined using the parameters -IncludePivotTable, -PivotColumns -PivotRows,
        -PivotData, -PivotFilter, -PivotTotals, -PivotDataToColumn, -IncludePivotChart and -ChartType.
        Its -PivotTableDefintion paramater allows multiple PivotTables to be defined, with additional parameters.
        New-PivotTableDefinition is a convenient way to build these definitions.
      .Example
        >
        PS>   $pt  = New-PivotTableDefinition -PivotTableName "PT1" -SourceWorkSheet "Sheet1" -PivotRows "Status"  -PivotData @{Status='Count' } -PivotFilter 'StartType' -IncludePivotChart  -ChartType BarClustered3D
        $Pt += New-PivotTableDefinition -PivotTableName "PT2" -SourceWorkSheet "Sheet2" -PivotRows "Company" -PivotData @{Company='Count'} -IncludePivotChart  -ChartType PieExploded3D  -ShowPercent -ChartTitle "Breakdown of processes by company"
        Get-Service | Select-Object    -Property Status,Name,DisplayName,StartType | Export-Excel -Path .\test.xlsx -AutoSize
        Get-Process | Select-Object    -Property Name,Company,Handles,CPU,VM       | Export-Excel -Path .\test.xlsx -AutoSize -WorksheetName 'sheet2'
        $excel = Export-Excel -Path .\test.xlsx -PivotTableDefinition $pt -Show

        This is a re-work of one of the examples in Export-Excel - instead of writing out the pivot definition hash table it is built by calling New-PivotTableDefinition.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',Justification='Does not change system State')]
    param(
        [Parameter(Mandatory)]
        [Alias("PivtoTableName")]#Previous typo - use alias to avoid breaking scripts
        $PivotTableName,
        #Worksheet where the data is found
        $SourceWorkSheet,
        #Address range in the worksheet e.g "A10:F20" - the first row must be column names: if not specified the whole sheet will be used/
        $SourceRange,
        #Fields to set as rows in the PivotTable
        $PivotRows,
        #A hash table in form "FieldName"="Function", where function is one of
        #Average, Count, CountNums, Max, Min, Product, None, StdDev, StdDevP, Sum, Var, VarP
        [hashtable]$PivotData,
        #Fields to set as columns in the PivotTable
        $PivotColumns,
        #Fields to use to filter in the PivotTable
        $PivotFilter,
        #If there are multiple datasets in a PivotTable, by default they are shown seperatate rows under the given row heading; this switch makes them seperate columns.
        [Switch]$PivotDataToColumn,
        #By default PivotTables have Totals for each Row (on the right) and for each column at the bottom. This allows just one or neither to be selected.
        #Define whther totals should be added to rows, columns neither, or both (the default is both)
        [ValidateSet("Both", "Columns", "Rows", "None")]
        [String]$PivotTotals = "Both",
        #Included for compatibility - equivalent to -PivotTotals "None"
        [Switch]$NoTotalsInPivot,
        #The name of a row field which should be grouped by parts of the date/time (ignored if GroupDateRow is not specified)
        [String]$GroupDateRow,
        #The Part(s) of the date to use in the grouping (ignored if GroupDateRow is not specified)
        [OfficeOpenXml.Table.PivotTable.eDateGroupBy[]]$GroupDatePart,
        #The name of a row field which should be grouped by Number (e.g 0-99, 100-199, 200-299 )
        [String]$GroupNumericRow,
        #The starting point for grouping
        [double]$GroupNumericMin = 0 ,
        #The endpoint for grouping
        [double]$GroupNumericMax = [Double]::MaxValue  ,
        #The interval for grouping
        [double]$GroupNumericInterval = 100  ,
        #Number format to apply to the data cells in the PivotTable
        [string]$PivotNumberFormat,
        #Apply a table style to the PivotTable
        [OfficeOpenXml.Table.TableStyles]$PivotTableStyle,
        #Use a chart definition instead of specifying chart settings one by one
        [Parameter(ParameterSetName = 'ChartbyDef', Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        $PivotChartDefinition,
        #If specified a chart Will be included.
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Switch]$IncludePivotChart,
        #Optional title for the pivot chart, by default the title omitted.
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [String]$ChartTitle,
        #Height of the chart in Pixels (400 by default)
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [int]$ChartHeight = 400 ,
        #Width of the chart in Pixels (600 by default)
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [int]$ChartWidth = 600,
        #Cell position of the top left corner of the chart, there will be this number of rows above the top edge of the chart (default is 0, chart starts at top edge of row 1).
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Int]$ChartRow = 0 ,
        #Cell position of the top left corner of the chart, there will be this number of cells to the left of the chart (default is 4, chart starts at left edge of column E)
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Int]$ChartColumn = 4,
        #Vertical offset of the chart from the cell corner.
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Int]$ChartRowOffSetPixels = 0 ,
        #Horizontal offset of the chart from the cell corner.
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Int]$ChartColumnOffSetPixels = 0,
        #Type of chart
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = 'Pie',
        #If specified hides the chart legend
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Switch]$NoLegend,
        #if specified attaches the category to slices in a pie chart : not supported on all chart types, this may give errors if applied to an unsupported type.
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Switch]$ShowCategory,
        #If specified attaches percentages to slices in a pie chart.
        [Parameter(ParameterSetName = 'ChartbyParams')]
        [Switch]$ShowPercent,
        #If there is already content in the workbook the sheet with the PivotTable will not be active UNLESS Activate is specified
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
    $parameters.Remove('PivotTableName')
    if ($PivotChartDefinition) {
        $parameters.PivotChartDefinition.XRange = $null
        $parameters.PivotChartDefinition.YRange = $null
        $parameters.PivotChartDefinition.SeriesHeader = $null
    }
    @{$PivotTableName = $parameters}
}
