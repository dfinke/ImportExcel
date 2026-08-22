function Invoke-ExcelReportPlan {
    <#
    .SYNOPSIS
        Executes an ImportExcel AI report plan.

    .DESCRIPTION
        Takes a plan from New-ExcelReportPlan, reads the source data, and creates an
        Excel workbook with summary sheets, formatted data tables, charts, pivot tables,
        number formats, and simple conditional formatting.

    .PARAMETER Plan
        Plan object, JSON string, or path to a JSON plan file.

    .PARAMETER Path
        Output .xlsx path.

    .PARAMETER SourcePath
        Source data path. If omitted, uses Plan.SourcePath or DatasetSummary.Path.

    .PARAMETER DatasetSummary
        Optional precomputed summary from Get-ExcelDatasetSummary.

    .PARAMETER Force
        Overwrites an existing output workbook. The source workbook is never overwritten.

    .PARAMETER Show
        Opens the workbook after it is created.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $Plan,

        [Parameter(Mandatory)]
        [string]$Path,

        [string]$SourcePath,

        $DatasetSummary,

        [string[]]$WorksheetName,

        [char]$Delimiter = ',',

        [switch]$Force,

        [switch]$Show
    )

    process {
        if ($Plan -is [string]) {
            if (Test-Path -LiteralPath $Plan) {
                $Plan = Get-Content -LiteralPath $Plan -Raw | ConvertFrom-Json
            }
            else {
                $Plan = $Plan | ConvertFrom-Json
            }
        }

        $planHash = ConvertTo-ExcelAiHashtable -InputObject $Plan
        if (-not $SourcePath) { $SourcePath = $planHash.SourcePath }
        if (-not $SourcePath -and $DatasetSummary) {
            $summaryHash = ConvertTo-ExcelAiHashtable -InputObject $DatasetSummary
            $SourcePath = $summaryHash.Path
        }
        if (-not $SourcePath) {
            throw "SourcePath is required when the plan does not contain SourcePath."
        }

        $resolvedSourcePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SourcePath)
        $resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)

        if ([string]::Equals($resolvedSourcePath, $resolvedOutputPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "The output path cannot be the same as the source path. Choose a different -Path."
        }

        if ((Test-Path -LiteralPath $resolvedOutputPath) -and -not $Force) {
            throw "Output workbook already exists: $resolvedOutputPath. Use -Force to overwrite it."
        }

        if (-not $DatasetSummary) {
            $DatasetSummary = Get-ExcelDatasetSummary -Path $resolvedSourcePath -WorksheetName $WorksheetName -Delimiter $Delimiter
        }
        $summaryHash = ConvertTo-ExcelAiHashtable -InputObject $DatasetSummary

        $sources = @(Get-ExcelAiSourceData -Path $resolvedSourcePath -WorksheetName $WorksheetName -Delimiter $Delimiter)
        $sourceByName = @{}
        foreach ($source in $sources) {
            $sourceByName[[string]$source.Name] = $source
        }

        if ((Test-Path -LiteralPath $resolvedOutputPath) -and $Force) {
            Remove-Item -LiteralPath $resolvedOutputPath -Force
        }

        $outputDirectory = Split-Path -Path $resolvedOutputPath -Parent
        if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
            $null = New-Item -Path $outputDirectory -ItemType Directory -Force
        }

        if (-not $PSCmdlet.ShouldProcess($resolvedOutputPath, 'Create Excel report workbook')) {
            return
        }

        $summaryPlan = $planHash.Summary
        $summarySheetName = if ($summaryPlan.SheetName) { Get-ExcelAiSafeName -Name $summaryPlan.SheetName } else { 'Summary' }
        $summaryTitle = if ($summaryPlan.Title) { [string]$summaryPlan.Title } elseif ($planHash.WorkbookTitle) { [string]$planHash.WorkbookTitle } else { 'Excel Report' }

        $summaryRows = @()
        foreach ($source in (Get-ExcelAiArray $summaryHash.Sources)) {
            $columns = @(Get-ExcelAiArray $source.ColumnSummaries)
            $summaryRows += [pscustomobject][ordered]@{
                Source = $source.Name
                Kind = $source.Kind
                Rows = $source.RowCount
                Columns = $source.ColumnCount
                Measures = @($columns | Where-Object { $_.Role -eq 'Measure' }).Count
                Dimensions = @($columns | Where-Object { $_.Role -eq 'Dimension' }).Count
                DateColumns = @($columns | Where-Object { $_.Role -eq 'Time' }).Count
                MissingValues = ($columns | Measure-Object -Property MissingCount -Sum).Sum
            }
        }

        if ($summaryRows.Count -eq 0) {
            $summaryRows = @([pscustomobject][ordered]@{
                Source = [System.IO.Path]::GetFileName($resolvedSourcePath)
                Kind = 'Unknown'
                Rows = 0
                Columns = 0
                Measures = 0
                Dimensions = 0
                DateColumns = 0
                MissingValues = 0
            })
        }

        $columnProfileRows = @()
        foreach ($source in (Get-ExcelAiArray $summaryHash.Sources)) {
            foreach ($column in (Get-ExcelAiArray $source.ColumnSummaries)) {
                $columnProfileRows += [pscustomobject][ordered]@{
                    Source = $source.Name
                    Column = $column.Name
                    InferredType = $column.InferredType
                    Role = $column.Role
                    NonEmpty = $column.NonEmptyCount
                    Missing = $column.MissingCount
                    Distinct = $column.DistinctCount
                    SuggestedFormat = $column.SuggestedNumberFormat
                }
            }
        }

        $usedSheetNames = @{}
        $usedSheetNames[$summarySheetName.ToLowerInvariant()] = $true
        $usedSheetNames['column profile'] = $true
        $tableSheetBySource = @{}
        $pivotDefinitionsBySource = @{}
        $sourceSummaryByName = @{}
        foreach ($sourceSummary in (Get-ExcelAiArray $summaryHash.Sources)) {
            $sourceSummaryByName[[string]$sourceSummary.Name] = $sourceSummary
        }

        $tables = @(Get-ExcelAiArray $planHash.Tables)
        if ($tables.Count -eq 0) {
            $tables = @((Get-ExcelAiDefaultReportPlan -DatasetSummary $DatasetSummary -Prompt $planHash.Prompt).Tables)
        }

        foreach ($table in $tables) {
            $sourceName = [string]$table.SourceName
            if (-not $sourceByName.ContainsKey($sourceName)) {
                Write-Warning "Skipping table '$($table.SheetName)': source '$sourceName' was not found."
                continue
            }

            $source = $sourceByName[$sourceName]
            $data = @($source.Data)
            if ($data.Count -eq 0) {
                Write-Warning "Skipping table '$($table.SheetName)': source '$sourceName' has no rows."
                continue
            }

            $sheetName = if ($table.SheetName) { [string]$table.SheetName } else { $sourceName }
            $sheetName = Get-ExcelAiUniqueSheetName -Name $sheetName -UsedNames $usedSheetNames
            $tableSheetBySource[$sourceName] = $sheetName
            $excelTableName = Get-ExcelAiTableName -Name $table.TableName -DefaultName ($sourceName + '_Data')

            $chartDefinitions = @()
            foreach ($chart in (Get-ExcelAiArray $planHash.Charts | Where-Object { [string]$_.SourceName -eq $sourceName })) {
                $chartParams = @{
                    Title = if ($chart.Title) { [string]$chart.Title } else { 'Chart' }
                    ChartType = if ($chart.ChartType) { [string]$chart.ChartType } else { 'ColumnClustered' }
                    XRange = [string]$chart.XColumn
                    YRange = @(Get-ExcelAiArray $chart.YColumn)
                    Width = if ($chart.Width) { [int]$chart.Width } else { 640 }
                    Height = if ($chart.Height) { [int]$chart.Height } else { 360 }
                    Row = if ($chart.Row) { [int]$chart.Row } else { 1 }
                    Column = if ($chart.Column) { [int]$chart.Column } else { 7 }
                }
                if ($chart.NoLegend) { $chartParams.NoLegend = $true }
                if ($chart.ShowCategory) { $chartParams.ShowCategory = $true }
                if ($chart.ShowPercent) { $chartParams.ShowPercent = $true }
                $chartDefinitions += New-ExcelChartDefinition @chartParams
            }

            $pivotDefinitions = @{}
            foreach ($pivot in (Get-ExcelAiArray $planHash.Pivots | Where-Object { [string]$_.SourceName -eq $sourceName })) {
                $pivotName = Get-ExcelAiTableName -Name $pivot.PivotTableName -DefaultName 'PivotSummary'
                if ($pivotName -eq $excelTableName) {
                    $pivotName = "${pivotName}_pivot"
                }
                $pivotName = Get-ExcelAiUniqueSheetName -Name $pivotName -UsedNames $usedSheetNames
                $pivotDefinition = @{
                    PivotRows = @(Get-ExcelAiArray $pivot.PivotRows)
                    PivotData = ConvertTo-ExcelAiHashtable -InputObject $pivot.PivotData
                    PivotTableStyle = if ($pivot.PivotTableStyle) { [string]$pivot.PivotTableStyle } else { 'Medium9' }
                }
                if ($pivot.PivotColumns) { $pivotDefinition.PivotColumns = @(Get-ExcelAiArray $pivot.PivotColumns) }
                if ($pivot.PivotFilter) { $pivotDefinition.PivotFilter = @(Get-ExcelAiArray $pivot.PivotFilter) }
                if ($pivot.IncludePivotChart) { $pivotDefinition.IncludePivotChart = $true }
                if ($pivot.ChartType) { $pivotDefinition.ChartType = [string]$pivot.ChartType }
                if ($pivot.ChartTitle) { $pivotDefinition.ChartTitle = [string]$pivot.ChartTitle }
                if ($pivot.NoLegend) { $pivotDefinition.NoLegend = $true }
                if ($pivot.ShowCategory) { $pivotDefinition.ShowCategory = $true }
                if ($pivot.ShowPercent) { $pivotDefinition.ShowPercent = $true }
                $pivotDefinitions[$pivotName] = $pivotDefinition
            }
            if ($pivotDefinitions.Count -gt 0) {
                $pivotDefinitionsBySource[$sourceName] = $pivotDefinitions
            }

            $exportParams = @{
                Path = $resolvedOutputPath
                WorksheetName = $sheetName
                TableName = $excelTableName
                TableStyle = if ($table.TableStyle) { [string]$table.TableStyle } else { 'Medium6' }
                AutoNameRange = $true
            }
            if ($table.AutoSize -ne $false) { $exportParams.AutoSize = $true }
            if ($table.AutoFilter -ne $false) { $exportParams.AutoFilter = $true }
            if ($table.BoldTopRow -ne $false) { $exportParams.BoldTopRow = $true }
            if ($table.FreezeTopRow -ne $false) { $exportParams.FreezeTopRow = $true }
            if ($chartDefinitions.Count -gt 0) { $exportParams.ExcelChartDefinition = $chartDefinitions }

            $data | Export-Excel @exportParams
        }

        foreach ($analysisSheet in (Get-ExcelAiArray $planHash.AnalysisSheets)) {
            $sourceName = [string]$analysisSheet.SourceName
            if (-not $sourceByName.ContainsKey($sourceName) -or -not $sourceSummaryByName.ContainsKey($sourceName)) {
                Write-Warning "Skipping analysis sheet '$($analysisSheet.SheetName)': source '$sourceName' was not found."
                continue
            }

            $source = $sourceByName[$sourceName]
            $sourceSummary = $sourceSummaryByName[$sourceName]
            $sheetName = Get-ExcelAiUniqueSheetName -Name $analysisSheet.SheetName -UsedNames $usedSheetNames
            $title = if ($analysisSheet.Title) { [string]$analysisSheet.Title } else { $sheetName }
            $tableStyle = if ($analysisSheet.TableStyle) { [string]$analysisSheet.TableStyle } else { 'Medium4' }
            $tableName = Get-ExcelAiTableName -Name ($sheetName + '_Table') -DefaultName 'AnalysisTable'

            $analysisRows = @()
            switch ([string]$analysisSheet.Type) {
                'ExecutiveDashboard' {
                    $analysisRows = @(New-ExcelAiExecutiveRows -Rows @($source.Data) -SourceSummary $sourceSummary)
                }
                'DataScience' {
                    $analysisRows = @(New-ExcelAiDataScienceRows -Rows @($source.Data) -SourceSummary $sourceSummary)
                }
                'CorrelationMatrix' {
                    $analysisRows = @(New-ExcelAiCorrelationRows -Rows @($source.Data) -SourceSummary $sourceSummary)
                }
                default {
                    Write-Warning "Skipping analysis sheet '$sheetName': unsupported type '$($analysisSheet.Type)'."
                    continue
                }
            }

            if ($analysisRows.Count -eq 0) {
                Write-Warning "Skipping analysis sheet '$sheetName': no rows were generated."
                continue
            }

            $analysisRows | Export-Excel -Path $resolvedOutputPath -WorksheetName $sheetName -Title $title -TitleBold -TitleSize 16 -AutoSize -BoldTopRow -TableName $tableName -TableStyle $tableStyle
        }

        $summaryRows | Export-Excel -Path $resolvedOutputPath -WorksheetName $summarySheetName -Title $summaryTitle -TitleBold -TitleSize 18 -AutoSize -BoldTopRow -TableName 'DatasetProfile' -TableStyle Medium2

        if ($columnProfileRows.Count -gt 0) {
            $columnProfileRows | Export-Excel -Path $resolvedOutputPath -WorksheetName 'Column Profile' -AutoSize -BoldTopRow -TableName 'ColumnProfile' -TableStyle Light9
        }

        $excel = Open-ExcelPackage -Path $resolvedOutputPath
        try {
            foreach ($worksheet in $excel.Workbook.Worksheets) {
                $worksheet.View.ShowGridLines = $false
            }

            foreach ($table in $tables) {
                $sourceName = [string]$table.SourceName
                if (-not $tableSheetBySource.ContainsKey($sourceName)) { continue }
                $sheetName = $tableSheetBySource[$sourceName]
                $worksheet = $excel.Workbook.Worksheets[$sheetName]
                if (-not $worksheet -or -not $worksheet.Dimension) { continue }

                $formats = ConvertTo-ExcelAiHashtable -InputObject $table.NumberFormats
                if ($formats) {
                    foreach ($columnName in $formats.Keys) {
                        $columnIndex = Get-ExcelAiColumnIndex -Worksheet $worksheet -HeaderName $columnName
                        if ($columnIndex -gt 0 -and $worksheet.Dimension.End.Row -gt 1) {
                            Set-ExcelRange -Address $worksheet.Cells[2, $columnIndex, $worksheet.Dimension.End.Row, $columnIndex] -NumberFormat $formats[$columnName]
                        }
                    }
                }
            }

            foreach ($sourceName in $pivotDefinitionsBySource.Keys) {
                if (-not $tableSheetBySource.ContainsKey($sourceName)) { continue }

                $sheetName = $tableSheetBySource[$sourceName]
                $worksheet = $excel.Workbook.Worksheets[$sheetName]
                if (-not $worksheet -or -not $worksheet.Dimension) { continue }

                $moveAfterSheetName = $sheetName
                foreach ($pivotItem in $pivotDefinitionsBySource[$sourceName].GetEnumerator()) {
                    $params = @{} + $pivotItem.Value
                    $params.SourceWorksheet = $worksheet
                    $params.SourceRange = $worksheet.Dimension.Address

                    try {
                        Add-PivotTable -ExcelPackage $excel -PivotTableName $pivotItem.Key @params
                        if ($excel.Workbook.Worksheets[$pivotItem.Key] -and $excel.Workbook.Worksheets[$moveAfterSheetName]) {
                            $excel.Workbook.Worksheets.MoveAfter($pivotItem.Key, $moveAfterSheetName)
                            $moveAfterSheetName = $pivotItem.Key
                        }
                    }
                    catch {
                        Write-Warning "Skipping pivot table '$($pivotItem.Key)' for source '$sourceName': $_"
                    }
                }
            }

            foreach ($format in (Get-ExcelAiArray $planHash.ConditionalFormats)) {
                $sourceName = [string]$format.SourceName
                if (-not $tableSheetBySource.ContainsKey($sourceName)) { continue }
                $sheetName = if ($format.SheetName) { [string]$format.SheetName } else { $tableSheetBySource[$sourceName] }
                if (-not $excel.Workbook.Worksheets[$sheetName]) { $sheetName = $tableSheetBySource[$sourceName] }
                $worksheet = $excel.Workbook.Worksheets[$sheetName]
                if (-not $worksheet -or -not $worksheet.Dimension) { continue }

                $columnIndex = Get-ExcelAiColumnIndex -Worksheet $worksheet -HeaderName ([string]$format.Column)
                if ($columnIndex -le 0 -or $worksheet.Dimension.End.Row -le 1) { continue }

                $columnLetter = (Get-ExcelColumnName -ColumnNumber $columnIndex).ColumnName
                $address = "$columnLetter`2:$columnLetter$($worksheet.Dimension.End.Row)"
                if ([string]$format.Type -eq 'DataBar') {
                    $color = if ($format.Color) { [string]$format.Color } else { 'SteelBlue' }
                    Add-ConditionalFormatting -Worksheet $worksheet -Address $address -DataBarColor $color | Out-Null
                }
                elseif ([string]$format.Type -in 'ThreeColorScale', 'TwoColorScale') {
                    Add-ConditionalFormatting -Worksheet $worksheet -Address $address -RuleType ([string]$format.Type) | Out-Null
                }
                elseif ([string]$format.Type -in 'Top', 'Bottom') {
                    $params = @{
                        Worksheet = $worksheet
                        Address = $address
                        RuleType = [string]$format.Type
                        ConditionValue = if ($format.Rank) { [int]$format.Rank } else { 10 }
                        BackgroundColor = if ($format.BackgroundColor) { [string]$format.BackgroundColor } elseif ([string]$format.Type -eq 'Top') { 'LightGreen' } else { 'LightPink' }
                        ForegroundColor = if ($format.FontColor) { [string]$format.FontColor } elseif ([string]$format.Type -eq 'Top') { 'DarkGreen' } else { 'DarkRed' }
                    }
                    if ($format.Bold) { $params.Bold = $true }
                    Add-ConditionalFormatting @params | Out-Null
                }
            }
        }
        finally {
            Close-ExcelPackage -ExcelPackage $excel -Show:$Show
        }

        [pscustomobject][ordered]@{
            Path = $resolvedOutputPath
            SourcePath = $resolvedSourcePath
            Plan = $Plan
            DatasetSummary = $DatasetSummary
            Worksheets = @(Open-ExcelPackage -Path $resolvedOutputPath | ForEach-Object {
                try { @($_.Workbook.Worksheets | Select-Object -ExpandProperty Name) }
                finally { Close-ExcelPackage -ExcelPackage $_ -NoSave }
            })
        }
    }
}
