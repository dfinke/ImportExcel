function ConvertTo-ExcelAiHashtable {
    param(
        [Parameter(ValueFromPipeline)]
        $InputObject
    )

    process {
        if ($null -eq $InputObject) { return $null }

        if ($InputObject -is [System.Collections.IDictionary]) {
            $hash = [ordered]@{}
            foreach ($key in $InputObject.Keys) {
                $hash[$key] = ConvertTo-ExcelAiHashtable -InputObject $InputObject[$key]
            }
            return $hash
        }

        if ($InputObject -is [string]) { return $InputObject }

        if ($InputObject -is [System.Collections.IEnumerable] -and
            $InputObject -isnot [System.Management.Automation.PSCustomObject]) {
            $items = @()
            foreach ($item in $InputObject) {
                $items += ConvertTo-ExcelAiHashtable -InputObject $item
            }
            return ,$items
        }

        if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
            $hash = [ordered]@{}
            foreach ($property in $InputObject.PSObject.Properties) {
                $hash[$property.Name] = ConvertTo-ExcelAiHashtable -InputObject $property.Value
            }
            return $hash
        }

        return $InputObject
    }
}

function Get-ExcelAiArray {
    param($Value)

    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) { return ,$Value }
    if ($Value -is [System.Collections.IDictionary]) { return ,$Value }
    if ($Value -is [System.Collections.IEnumerable]) { return @($Value) }
    return @($Value)
}

function Get-ExcelAiSafeName {
    param(
        [string]$Name,
        [int]$MaxLength = 31,
        [string]$DefaultName = 'Sheet'
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = $DefaultName
    }

    $safe = $Name -replace '[\\/\?\*\[\]:]', '_'
    $safe = $safe.Trim("'")
    $safe = $safe.Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = $DefaultName
    }

    if ($safe.Length -gt $MaxLength) {
        $safe = $safe.Substring(0, $MaxLength)
    }

    return $safe
}

function Get-ExcelAiUniqueSheetName {
    param(
        [string]$Name,
        [hashtable]$UsedNames
    )

    $baseName = Get-ExcelAiSafeName -Name $Name
    $candidate = $baseName
    $index = 1

    while ($UsedNames.ContainsKey($candidate.ToLowerInvariant())) {
        $suffix = "_$index"
        $maxBaseLength = 31 - $suffix.Length
        if ($baseName.Length -gt $maxBaseLength) {
            $candidate = $baseName.Substring(0, $maxBaseLength) + $suffix
        }
        else {
            $candidate = $baseName + $suffix
        }
        $index++
    }

    $UsedNames[$candidate.ToLowerInvariant()] = $true
    return $candidate
}

function Get-ExcelAiTableName {
    param(
        [string]$Name,
        [string]$DefaultName = 'Table1'
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = $DefaultName
    }

    $safe = $Name -replace '[^A-Za-z0-9_]', '_'
    if ($safe -notmatch '^[A-Za-z_]') {
        $safe = "T_$safe"
    }
    if ($safe.Length -gt 240) {
        $safe = $safe.Substring(0, 240)
    }

    return $safe
}

function Get-ExcelAiSourceData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [string[]]$WorksheetName,

        [char]$Delimiter = ','
    )

    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if (-not (Test-Path -LiteralPath $resolvedPath)) {
        throw "Path not found: $Path"
    }

    $extension = [System.IO.Path]::GetExtension($resolvedPath).ToLowerInvariant()
    $sources = @()

    switch ($extension) {
        '.csv' {
            $data = @(Import-Csv -LiteralPath $resolvedPath -Delimiter $Delimiter)
            $sources += [pscustomobject]@{
                Name = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath)
                Kind = 'Csv'
                Path = $resolvedPath
                Data = $data
            }
        }
        '.tsv' {
            $data = @(Import-Csv -LiteralPath $resolvedPath -Delimiter "`t")
            $sources += [pscustomobject]@{
                Name = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath)
                Kind = 'Tsv'
                Path = $resolvedPath
                Data = $data
            }
        }
        { $_ -in '.xlsx', '.xlsm', '.xltx', '.xltm' } {
            $sheetNames = @()
            if ($WorksheetName) {
                $sheetNames = $WorksheetName
            }
            else {
                $sheetNames = @(Get-ExcelSheetInfo -Path $resolvedPath | Where-Object { -not $_.Hidden } | Select-Object -ExpandProperty Name)
            }

            foreach ($sheetName in $sheetNames) {
                $data = @(Import-Excel -Path $resolvedPath -WorksheetName $sheetName)
                $sources += [pscustomobject]@{
                    Name = $sheetName
                    Kind = 'Worksheet'
                    Path = $resolvedPath
                    Data = $data
                }
            }
        }
        default {
            throw "Unsupported file type '$extension'. Supported inputs are .csv, .tsv, .xlsx, .xlsm, .xltx, and .xltm."
        }
    }

    return $sources
}

function Test-ExcelAiEmptyValue {
    param($Value)

    if ($null -eq $Value) { return $true }
    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return $true }
    return $false
}

function Test-ExcelAiLong {
    param([string]$Value)

    $number = 0L
    return [long]::TryParse(
        $Value,
        [System.Globalization.NumberStyles]::Integer,
        [System.Globalization.CultureInfo]::CurrentCulture,
        [ref]$number
    )
}

function Test-ExcelAiDouble {
    param([string]$Value)

    $number = 0.0
    $styles = [System.Globalization.NumberStyles]::Float -bor
        [System.Globalization.NumberStyles]::AllowThousands -bor
        [System.Globalization.NumberStyles]::AllowCurrencySymbol -bor
        [System.Globalization.NumberStyles]::AllowParentheses

    return [double]::TryParse(
        $Value,
        $styles,
        [System.Globalization.CultureInfo]::CurrentCulture,
        [ref]$number
    )
}

function ConvertTo-ExcelAiDouble {
    param($Value)

    if (Test-ExcelAiEmptyValue $Value) { return $null }

    $number = 0.0
    $styles = [System.Globalization.NumberStyles]::Float -bor
        [System.Globalization.NumberStyles]::AllowThousands -bor
        [System.Globalization.NumberStyles]::AllowCurrencySymbol -bor
        [System.Globalization.NumberStyles]::AllowParentheses

    if ([double]::TryParse(
            [string]$Value,
            $styles,
            [System.Globalization.CultureInfo]::CurrentCulture,
            [ref]$number
        )) {
        return $number
    }

    return $null
}

function Test-ExcelAiDate {
    param([string]$Value)

    if ($Value -match '^\s*[-+]?\d+(\.\d+)?\s*$') { return $false }

    $date = [datetime]::MinValue
    return [datetime]::TryParse(
        $Value,
        [System.Globalization.CultureInfo]::CurrentCulture,
        [System.Globalization.DateTimeStyles]::None,
        [ref]$date
    )
}

function Test-ExcelAiBoolean {
    param([string]$Value)

    $bool = $false
    if ([bool]::TryParse($Value, [ref]$bool)) { return $true }
    return $Value -match '^(0|1|yes|no|y|n)$'
}

function Get-ExcelAiColumnSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,

        [int]$MaxDistinctValues = 10
    )

    if ($Rows.Count -eq 0) {
        return @()
    }

    $propertyNames = @($Rows[0].PSObject.Properties | Select-Object -ExpandProperty Name)
    $summaries = @()

    foreach ($name in $propertyNames) {
        $values = @($Rows | ForEach-Object { $_.PSObject.Properties[$name].Value })
        $nonEmptyValues = @($values | Where-Object { -not (Test-ExcelAiEmptyValue $_) })
        $distinctValues = @($nonEmptyValues | ForEach-Object { [string]$_ } | Sort-Object -Unique)

        $inferredType = 'Empty'
        if ($nonEmptyValues.Count -gt 0) {
            $allInteger = $true
            $allNumber = $true
            $allDate = $true
            $allBoolean = $true

            foreach ($value in $nonEmptyValues) {
                $text = [string]$value
                if (-not (Test-ExcelAiLong -Value $text)) { $allInteger = $false }
                if (-not (Test-ExcelAiDouble -Value $text)) { $allNumber = $false }
                if (-not (Test-ExcelAiDate -Value $text)) { $allDate = $false }
                if (-not (Test-ExcelAiBoolean -Value $text)) { $allBoolean = $false }
            }

            if ($allBoolean) { $inferredType = 'Boolean' }
            elseif ($allInteger) { $inferredType = 'Integer' }
            elseif ($allNumber) { $inferredType = 'Number' }
            elseif ($allDate) { $inferredType = 'DateTime' }
            else { $inferredType = 'Text' }
        }

        $role = 'Dimension'
        if ($inferredType -in 'Integer', 'Number') { $role = 'Measure' }
        elseif ($inferredType -eq 'DateTime') { $role = 'Time' }

        $suggestedFormat = $null
        if ($inferredType -in 'Integer', 'Number') {
            if ($name -match '(amount|cost|price|revenue|sales|total|profit|margin|dollar|budget)') {
                $suggestedFormat = 'Currency'
            }
            elseif ($name -match '(percent|percentage|pct|rate|ratio)') {
                $suggestedFormat = 'Percentage'
            }
            elseif ($inferredType -eq 'Integer') {
                $suggestedFormat = '#,##0'
            }
            else {
                $suggestedFormat = '#,##0.00'
            }
        }
        elseif ($inferredType -eq 'DateTime') {
            $suggestedFormat = 'Short Date'
        }

        $summaries += [pscustomobject][ordered]@{
            Name = $name
            InferredType = $inferredType
            Role = $role
            NonEmptyCount = $nonEmptyValues.Count
            MissingCount = $values.Count - $nonEmptyValues.Count
            DistinctCount = $distinctValues.Count
            Examples = @($distinctValues | Select-Object -First $MaxDistinctValues)
            SuggestedNumberFormat = $suggestedFormat
        }
    }

    return $summaries
}

function Get-ExcelAiPromptChartDefaults {
    [CmdletBinding()]
    param(
        [string]$Prompt
    )

    $normalizedPrompt = if ($Prompt) { $Prompt.ToLowerInvariant() } else { '' }

    $chartType = 'ColumnClustered'
    $noLegend = $true
    $showCategory = $false
    $showPercent = $false

    if ($normalizedPrompt -match '\b(pie|donut|doughnut)\b') {
        $chartType = if ($normalizedPrompt -match '\b(donut|doughnut)\b') { 'Doughnut' } else { 'Pie' }
        $noLegend = $false
        $showCategory = $true
        $showPercent = $true
    }
    elseif ($normalizedPrompt -match '\b(line|trend)\b') {
        $chartType = 'Line'
    }
    elseif ($normalizedPrompt -match '\b(bar|horizontal bar)\b') {
        $chartType = 'BarClustered'
    }
    elseif ($normalizedPrompt -match '\b(area)\b') {
        $chartType = 'Area'
    }

    [pscustomobject][ordered]@{
        ChartType = $chartType
        NoLegend = $noLegend
        ShowCategory = $showCategory
        ShowPercent = $showPercent
    }
}

function Get-ExcelAiPromptConditionalFormatSpecs {
    [CmdletBinding()]
    param(
        [string]$Prompt,

        [switch]$DefaultToDataBar
    )

    $normalizedPrompt = if ($Prompt) { $Prompt.ToLowerInvariant() } else { '' }
    $specs = @()

    $useDataBars = $normalizedPrompt -match '\bdata\s*bars?\b'
    $useColorScale = $normalizedPrompt -match '\b(colou?r\s*scale|heat\s*map|heatmap|gradient)\b'
    $useTop = $normalizedPrompt -match '\b(top|best|highest|largest)\b'
    $useBottom = $normalizedPrompt -match '\b(bottom|worst|lowest|smallest)\b'
    $hasExplicitFormatting = $useDataBars -or $useColorScale -or $useTop -or $useBottom

    if ($useDataBars -or ($DefaultToDataBar -and -not $hasExplicitFormatting)) {
        $specs += [pscustomobject][ordered]@{
            Type = 'DataBar'
            Color = 'SteelBlue'
        }
    }

    if ($useColorScale) {
        $specs += [pscustomobject][ordered]@{
            Type = 'ThreeColorScale'
        }
    }

    if ($useTop) {
        $rank = 10
        $rankMatch = [regex]::Match($normalizedPrompt, '\btop\s+(\d{1,3})\b')
        if ($rankMatch.Success) { $rank = [int]$rankMatch.Groups[1].Value }
        $specs += [pscustomobject][ordered]@{
            Type = 'Top'
            Rank = $rank
            BackgroundColor = 'LightGreen'
            FontColor = 'DarkGreen'
            Bold = $true
        }
    }

    if ($useBottom) {
        $rank = 10
        $rankMatch = [regex]::Match($normalizedPrompt, '\bbottom\s+(\d{1,3})\b')
        if ($rankMatch.Success) { $rank = [int]$rankMatch.Groups[1].Value }
        $specs += [pscustomobject][ordered]@{
            Type = 'Bottom'
            Rank = $rank
            BackgroundColor = 'LightPink'
            FontColor = 'DarkRed'
            Bold = $true
        }
    }

    return $specs
}

function Get-ExcelAiDefaultReportPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $DatasetSummary,

        [string]$Prompt
    )

    $summary = ConvertTo-ExcelAiHashtable -InputObject $DatasetSummary
    $chartDefaults = Get-ExcelAiPromptChartDefaults -Prompt $Prompt
    $conditionalFormatSpecs = @(Get-ExcelAiPromptConditionalFormatSpecs -Prompt $Prompt -DefaultToDataBar)
    $plan = [ordered]@{
        Version = '1.0'
        WorkbookTitle = if ($Prompt) { 'AI-Assisted Excel Report' } else { 'Excel Report' }
        Prompt = $Prompt
        SourcePath = $summary.Path
        Summary = [ordered]@{
            SheetName = 'Summary'
            Title = if ($Prompt) { 'AI-Assisted Excel Report' } else { 'Excel Report' }
            IncludeDatasetProfile = $true
        }
        Tables = @()
        Charts = @()
        Pivots = @()
        ConditionalFormats = @()
        AnalysisSheets = @()
    }

    foreach ($source in (Get-ExcelAiArray $summary.Sources)) {
        $sourceName = [string]$source.Name
        $sheetName = Get-ExcelAiSafeName -Name $sourceName -DefaultName 'Data'
        if ($sheetName -eq 'Summary') { $sheetName = 'Data' }

        $numberFormats = [ordered]@{}
        $columns = @(Get-ExcelAiArray $source.ColumnSummaries)
        foreach ($column in $columns) {
            if ($column.SuggestedNumberFormat) {
                $numberFormats[[string]$column.Name] = $column.SuggestedNumberFormat
            }
        }

        $plan.Tables += [ordered]@{
            SourceName = $sourceName
            SheetName = $sheetName
            TableName = Get-ExcelAiTableName -Name ($sourceName + '_Data') -DefaultName 'DataTable'
            TableStyle = 'Medium6'
            AutoSize = $true
            AutoFilter = $true
            BoldTopRow = $true
            FreezeTopRow = $true
            NumberFormats = $numberFormats
        }

        $measure = @($columns | Where-Object { $_.Role -eq 'Measure' } | Select-Object -First 1)
        $dimension = @($columns | Where-Object { $_.Role -eq 'Dimension' -and $_.DistinctCount -gt 1 -and $_.DistinctCount -le 30 } | Select-Object -First 1)
        if (-not $dimension) {
            $dimension = @($columns | Where-Object { $_.Role -eq 'Time' } | Select-Object -First 1)
        }

        if ($measure.Count -gt 0 -and $dimension.Count -gt 0) {
            $plan.Charts += [ordered]@{
                SourceName = $sourceName
                Title = "$($measure[0].Name) by $($dimension[0].Name)"
                ChartType = $chartDefaults.ChartType
                XColumn = [string]$dimension[0].Name
                YColumn = @([string]$measure[0].Name)
                Width = 640
                Height = 360
                Row = 1
                Column = 7
                NoLegend = $chartDefaults.NoLegend
                ShowCategory = $chartDefaults.ShowCategory
                ShowPercent = $chartDefaults.ShowPercent
            }

            $pivotData = [ordered]@{}
            $pivotData[[string]$measure[0].Name] = 'Sum'
            $pivotName = Get-ExcelAiTableName -Name ("Pivot_$($sourceName)_$($dimension[0].Name)") -DefaultName 'PivotSummary'
            $pivotName = Get-ExcelAiSafeName -Name $pivotName -MaxLength 31 -DefaultName 'PivotSummary'
            $plan.Pivots += [ordered]@{
                SourceName = $sourceName
                PivotTableName = $pivotName
                PivotRows = @([string]$dimension[0].Name)
                PivotData = $pivotData
                PivotTableStyle = 'Medium9'
                IncludePivotChart = $true
                ChartType = $chartDefaults.ChartType
                ChartTitle = "$($measure[0].Name) by $($dimension[0].Name)"
                NoLegend = $chartDefaults.NoLegend
                ShowCategory = $chartDefaults.ShowCategory
                ShowPercent = $chartDefaults.ShowPercent
            }

            foreach ($formatSpec in $conditionalFormatSpecs) {
                $conditionalFormat = [ordered]@{
                    SourceName = $sourceName
                    SheetName = $sheetName
                    Column = [string]$measure[0].Name
                    Type = [string]$formatSpec.Type
                }
                foreach ($propertyName in 'Color', 'Rank', 'BackgroundColor', 'FontColor', 'Bold') {
                    if ($null -ne $formatSpec.$propertyName) {
                        $conditionalFormat[$propertyName] = $formatSpec.$propertyName
                    }
                }
                $plan.ConditionalFormats += $conditionalFormat
            }
        }
    }

    $plan = Add-ExcelAiPromptAnalysisDefaults -Plan $plan -DatasetSummary $DatasetSummary -Prompt $Prompt

    return $plan
}

function Add-ExcelAiPromptAnalysisDefaults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Plan,

        [Parameter(Mandatory)]
        $DatasetSummary,

        [string]$Prompt
    )

    $planHash = ConvertTo-ExcelAiHashtable -InputObject $Plan
    if (-not $planHash.Contains('AnalysisSheets') -or $null -eq $planHash.AnalysisSheets) {
        $planHash.AnalysisSheets = @()
    }

    $existingKeys = @{}
    foreach ($sheet in (Get-ExcelAiArray $planHash.AnalysisSheets)) {
        $existingKeys["$($sheet.SourceName)|$($sheet.Type)"] = $true
    }

    $promptText = if ($Prompt) { $Prompt.ToLowerInvariant() } else { '' }
    $wantsDataScience = $promptText -match 'data scientist|statistic|statistical|correlation|distribution|outlier|regression|analy[sz]e'
    $wantsExecutive = $promptText -match 'executive|dashboard|kpi|leader|leadership|management|board|summary'

    if (-not $wantsDataScience -and -not $wantsExecutive) {
        return [pscustomobject]$planHash
    }

    $summary = ConvertTo-ExcelAiHashtable -InputObject $DatasetSummary
    foreach ($source in (Get-ExcelAiArray $summary.Sources)) {
        $sourceName = [string]$source.Name
        $columns = @(Get-ExcelAiArray $source.ColumnSummaries)
        $measureCount = @($columns | Where-Object { $_.Role -eq 'Measure' }).Count

        if ($wantsExecutive -and -not $existingKeys.ContainsKey("$sourceName|ExecutiveDashboard") -and $measureCount -gt 0) {
            $planHash.AnalysisSheets += [ordered]@{
                SourceName = $sourceName
                Type = 'ExecutiveDashboard'
                SheetName = 'Executive Dashboard'
                Title = 'Executive Dashboard'
                TableStyle = 'Medium4'
            }
        }

        if ($wantsDataScience -and -not $existingKeys.ContainsKey("$sourceName|DataScience") -and $measureCount -gt 0) {
            $planHash.AnalysisSheets += [ordered]@{
                SourceName = $sourceName
                Type = 'DataScience'
                SheetName = 'Statistical Analysis'
                Title = 'Statistical Analysis'
                TableStyle = 'Medium7'
            }
        }

        if ($wantsDataScience -and -not $existingKeys.ContainsKey("$sourceName|CorrelationMatrix") -and $measureCount -gt 1) {
            $planHash.AnalysisSheets += [ordered]@{
                SourceName = $sourceName
                Type = 'CorrelationMatrix'
                SheetName = 'Correlation Matrix'
                Title = 'Correlation Matrix'
                TableStyle = 'Light11'
            }
        }
    }

    return [pscustomobject]$planHash
}

function Get-ExcelAiMeasureColumns {
    param(
        [Parameter(Mandatory)]
        $SourceSummary
    )

    @(Get-ExcelAiArray $SourceSummary.ColumnSummaries | Where-Object { $_.Role -eq 'Measure' })
}

function Get-ExcelAiDimensionColumns {
    param(
        [Parameter(Mandatory)]
        $SourceSummary
    )

    @(Get-ExcelAiArray $SourceSummary.ColumnSummaries |
        Where-Object { $_.Role -eq 'Dimension' -and $_.DistinctCount -gt 1 -and $_.DistinctCount -le 50 })
}

function Get-ExcelAiNumericValues {
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,

        [Parameter(Mandatory)]
        [string]$ColumnName
    )

    @($Rows | ForEach-Object { ConvertTo-ExcelAiDouble $_.PSObject.Properties[$ColumnName].Value } | Where-Object { $null -ne $_ })
}

function Get-ExcelAiAverage {
    param([double[]]$Values)

    if (-not $Values -or $Values.Count -eq 0) { return $null }
    return ($Values | Measure-Object -Average).Average
}

function Get-ExcelAiMedian {
    param([double[]]$Values)

    if (-not $Values -or $Values.Count -eq 0) { return $null }
    $sorted = @($Values | Sort-Object)
    $middle = [int]($sorted.Count / 2)
    if ($sorted.Count % 2) {
        return $sorted[$middle]
    }
    return ($sorted[$middle - 1] + $sorted[$middle]) / 2
}

function Get-ExcelAiStandardDeviation {
    param([double[]]$Values)

    if (-not $Values -or $Values.Count -lt 2) { return $null }
    $average = Get-ExcelAiAverage -Values $Values
    $sumSquares = 0.0
    foreach ($value in $Values) {
        $sumSquares += [math]::Pow(($value - $average), 2)
    }
    return [math]::Sqrt($sumSquares / ($Values.Count - 1))
}

function Get-ExcelAiCorrelation {
    param(
        [double[]]$X,
        [double[]]$Y
    )

    if (-not $X -or -not $Y -or $X.Count -ne $Y.Count -or $X.Count -lt 2) { return $null }

    $avgX = Get-ExcelAiAverage -Values $X
    $avgY = Get-ExcelAiAverage -Values $Y
    $sumXY = 0.0
    $sumXX = 0.0
    $sumYY = 0.0

    for ($i = 0; $i -lt $X.Count; $i++) {
        $dx = $X[$i] - $avgX
        $dy = $Y[$i] - $avgY
        $sumXY += $dx * $dy
        $sumXX += $dx * $dx
        $sumYY += $dy * $dy
    }

    if ($sumXX -eq 0 -or $sumYY -eq 0) { return $null }
    return $sumXY / [math]::Sqrt($sumXX * $sumYY)
}

function New-ExcelAiDataScienceRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,

        [Parameter(Mandatory)]
        $SourceSummary
    )

    $result = @()
    foreach ($column in (Get-ExcelAiMeasureColumns -SourceSummary $SourceSummary)) {
        $values = [double[]](Get-ExcelAiNumericValues -Rows $Rows -ColumnName $column.Name)
        if ($values.Count -eq 0) { continue }

        $average = Get-ExcelAiAverage -Values $values
        $stdDev = Get-ExcelAiStandardDeviation -Values $values
        $min = ($values | Measure-Object -Minimum).Minimum
        $max = ($values | Measure-Object -Maximum).Maximum
        $note = if ($stdDev -and $average -ne 0 -and ([math]::Abs($stdDev / $average) -gt 0.5)) { 'High variation' } else { '' }

        $result += [pscustomobject][ordered]@{
            Column = $column.Name
            Count = $values.Count
            Missing = $column.MissingCount
            Sum = ($values | Measure-Object -Sum).Sum
            Mean = $average
            Median = Get-ExcelAiMedian -Values $values
            StandardDeviation = $stdDev
            Minimum = $min
            Maximum = $max
            Range = $max - $min
            Distinct = $column.DistinctCount
            Note = $note
        }
    }

    return $result
}

function New-ExcelAiExecutiveRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,

        [Parameter(Mandatory)]
        $SourceSummary
    )

    $result = @()
    $measures = @(Get-ExcelAiMeasureColumns -SourceSummary $SourceSummary)
    $dimensions = @(Get-ExcelAiDimensionColumns -SourceSummary $SourceSummary)

    foreach ($measure in $measures) {
        $values = [double[]](Get-ExcelAiNumericValues -Rows $Rows -ColumnName $measure.Name)
        if ($values.Count -eq 0) { continue }

        $result += [pscustomobject][ordered]@{
            Metric = "Total $($measure.Name)"
            Value = ($values | Measure-Object -Sum).Sum
            Detail = "$($Rows.Count) rows"
        }
        $result += [pscustomobject][ordered]@{
            Metric = "Average $($measure.Name)"
            Value = Get-ExcelAiAverage -Values $values
            Detail = 'Mean value'
        }
    }

    if ($measures.Count -gt 0 -and $dimensions.Count -gt 0) {
        $measureName = [string]$measures[0].Name
        foreach ($dimension in ($dimensions | Select-Object -First 2)) {
            $groups = @{}
            foreach ($row in $Rows) {
                $key = [string]$row.PSObject.Properties[$dimension.Name].Value
                $value = ConvertTo-ExcelAiDouble $row.PSObject.Properties[$measureName].Value
                if ([string]::IsNullOrWhiteSpace($key) -or $null -eq $value) { continue }
                if (-not $groups.ContainsKey($key)) { $groups[$key] = 0.0 }
                $groups[$key] += $value
            }

            $top = $groups.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1
            if ($top) {
                $result += [pscustomobject][ordered]@{
                    Metric = "Top $($dimension.Name) by $measureName"
                    Value = $top.Value
                    Detail = $top.Key
                }
            }
        }
    }

    return $result
}

function New-ExcelAiCorrelationRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,

        [Parameter(Mandatory)]
        $SourceSummary
    )

    $measures = @(Get-ExcelAiMeasureColumns -SourceSummary $SourceSummary)
    $result = @()

    foreach ($xColumn in $measures) {
        $row = [ordered]@{ Column = $xColumn.Name }
        foreach ($yColumn in $measures) {
            $xValues = New-Object System.Collections.Generic.List[double]
            $yValues = New-Object System.Collections.Generic.List[double]

            foreach ($dataRow in $Rows) {
                $x = ConvertTo-ExcelAiDouble $dataRow.PSObject.Properties[$xColumn.Name].Value
                $y = ConvertTo-ExcelAiDouble $dataRow.PSObject.Properties[$yColumn.Name].Value
                if ($null -ne $x -and $null -ne $y) {
                    $xValues.Add($x)
                    $yValues.Add($y)
                }
            }

            $row[[string]$yColumn.Name] = Get-ExcelAiCorrelation -X ([double[]]$xValues.ToArray()) -Y ([double[]]$yValues.ToArray())
        }
        $result += [pscustomobject]$row
    }

    return $result
}

function Import-ExcelAiPSAISuite {
    [CmdletBinding()]
    param(
        [string]$PSAISuitePath
    )

    if (Get-Command Invoke-ChatCompletion -ErrorAction SilentlyContinue) {
        return
    }

    $moduleRoot = Split-Path -Path $PSScriptRoot -Parent
    $repoRoot = Split-Path -Path $moduleRoot -Parent
    $candidatePaths = @()
    if ($PSAISuitePath) { $candidatePaths += $PSAISuitePath }
    if ($env:PSAISUITE_PATH) { $candidatePaths += $env:PSAISUITE_PATH }
    $candidatePaths += (Join-Path -Path $repoRoot -ChildPath 'psaisuite\PSAISuite.psd1')
    $candidatePaths += 'PSAISuite'

    foreach ($candidate in $candidatePaths) {
        try {
            if ($candidate -eq 'PSAISuite' -or (Test-Path -LiteralPath $candidate)) {
                Import-Module $candidate -ErrorAction Stop
                if (Get-Command Invoke-ChatCompletion -ErrorAction SilentlyContinue) {
                    return
                }
            }
        }
        catch {
            Write-Verbose "Could not import PSAISuite from '$candidate': $_"
        }
    }

    throw "PSAISuite was not found. Import PSAISuite first, set `$env:PSAISUITE_PATH, or pass -PSAISuitePath."
}

function ConvertFrom-ExcelAiJsonResponse {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Text
    )

    $candidate = $Text.Trim()
    if ($candidate -match '(?s)```(?:json)?\s*(.*?)\s*```') {
        $candidate = $Matches[1].Trim()
    }
    else {
        $firstBrace = $candidate.IndexOf('{')
        $lastBrace = $candidate.LastIndexOf('}')
        if ($firstBrace -ge 0 -and $lastBrace -gt $firstBrace) {
            $candidate = $candidate.Substring($firstBrace, $lastBrace - $firstBrace + 1)
        }
    }

    return $candidate | ConvertFrom-Json
}

function ConvertFrom-ExcelAiPowerShellResponse {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Text
    )

    $candidate = $Text.Trim()
    if ($candidate -match '(?s)```(?:powershell|pwsh|ps1)?\s*(.*?)\s*```') {
        $candidate = $Matches[1].Trim()
    }

    return $candidate
}

function ConvertTo-ExcelAiPowerShellLiteral {
    param($Value)

    if ($null -eq $Value) { return '$null' }
    return "'" + ([string]$Value -replace "'", "''") + "'"
}

function Test-ExcelAiPowerShellScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Script
    )

    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseInput($Script, [ref]$tokens, [ref]$errors) | Out-Null
    return @($errors)
}

function Test-ExcelAiPowerShellCommandUsage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Script
    )

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($Script, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) {
        return @()
    }

    $issues = @()
    $commandAsts = $ast.FindAll(
        { param($node) $node -is [System.Management.Automation.Language.CommandAst] },
        $true
    )

    foreach ($commandAst in $commandAsts) {
        $commandName = $commandAst.GetCommandName()
        if ([string]::IsNullOrWhiteSpace($commandName)) { continue }
        if ($commandName -match '^[\$\&]') { continue }

        $commandInfo = Get-Command -Name $commandName -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $commandInfo) {
            $issues += [pscustomobject][ordered]@{
                Command = $commandName
                Parameter = $null
                Message = "Command '$commandName' was not found."
            }
            continue
        }

        for ($elementIndex = 0; $elementIndex -lt $commandAst.CommandElements.Count; $elementIndex++) {
            $element = $commandAst.CommandElements[$elementIndex]
            if ($element -isnot [System.Management.Automation.Language.CommandParameterAst]) { continue }
            if ([string]::IsNullOrWhiteSpace($element.ParameterName)) { continue }

            $parameterName = $element.ParameterName
            if ($parameterName -in @('Verbose', 'Debug', 'ErrorAction', 'WarningAction', 'InformationAction', 'ErrorVariable', 'WarningVariable', 'InformationVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable', 'WhatIf', 'Confirm')) {
                continue
            }

            $exactMatches = @(
                foreach ($parameter in $commandInfo.Parameters.GetEnumerator()) {
                    if ($parameter.Key -eq $parameterName) {
                        $parameter.Key
                        continue
                    }

                    foreach ($alias in @($parameter.Value.Aliases)) {
                        if ($alias -eq $parameterName) {
                            $parameter.Key
                            break
                        }
                    }
                }
            ) | Select-Object -Unique
            $exactMatches = @($exactMatches)

            $matches = if ($exactMatches.Count -gt 0) {
                $exactMatches
            }
            else {
                @(
                foreach ($parameter in $commandInfo.Parameters.GetEnumerator()) {
                    if ($parameter.Key -like "$parameterName*") {
                        $parameter.Key
                        continue
                    }

                    foreach ($alias in @($parameter.Value.Aliases)) {
                        if ($alias -like "$parameterName*") {
                            $parameter.Key
                            break
                        }
                    }
                }
                ) | Select-Object -Unique
            }
            $matches = @($matches)
            if ($matches.Count -eq 0) {
                $issues += [pscustomobject][ordered]@{
                    Command = $commandName
                    Parameter = $parameterName
                    Message = "Command '$commandName' does not have a '-$parameterName' parameter."
                }
            }
            elseif ($matches.Count -gt 1) {
                $issues += [pscustomobject][ordered]@{
                    Command = $commandName
                    Parameter = $parameterName
                    Message = "Parameter '-$parameterName' is ambiguous for command '$commandName'."
                }
            }
            else {
                $resolvedParameterName = [string]$matches[0]
                $parameterMetadata = $commandInfo.Parameters[$resolvedParameterName]
                if (-not $parameterMetadata) { continue }

                $parameterType = $parameterMetadata.ParameterType
                $isSwitchParameter = $parameterType -eq [System.Management.Automation.SwitchParameter]
                if ($isSwitchParameter -or $null -ne $element.Argument) { continue }

                $nextElement = if (($elementIndex + 1) -lt $commandAst.CommandElements.Count) {
                    $commandAst.CommandElements[$elementIndex + 1]
                }
                else {
                    $null
                }

                if ($null -eq $nextElement -or $nextElement -is [System.Management.Automation.Language.CommandParameterAst]) {
                    $issues += [pscustomobject][ordered]@{
                        Command = $commandName
                        Parameter = $parameterName
                        Message = "Parameter '-$parameterName' for command '$commandName' requires an argument."
                    }
                }
            }
        }
    }

    $memberCallAsts = $ast.FindAll(
        { param($node) $node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] },
        $true
    )

    foreach ($memberCallAst in $memberCallAsts) {
        $memberName = $memberCallAst.Member.Extent.Text
        if ($memberName -eq 'SetColor') {
            $issues += [pscustomobject][ordered]@{
                Command = 'SetColor'
                Parameter = $null
                Message = "Direct EPPlus SetColor calls are not supported in generated scripts. Use ImportExcel color parameters such as -BackgroundColor, -FontColor, or -BorderColor."
            }
        }
    }

    return $issues
}

function New-ExcelAiDefaultReportScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $DatasetSummary,

        [Parameter(Mandatory)]
        [string]$SourcePath,

        [Parameter(Mandatory)]
        [string]$WorkbookPath,

        [string]$Prompt
    )

    $summary = ConvertTo-ExcelAiHashtable -InputObject $DatasetSummary
    $source = @(Get-ExcelAiArray $summary.Sources | Select-Object -First 1)[0]
    $columns = @(Get-ExcelAiArray $source.ColumnSummaries)
    $measure = @($columns | Where-Object { $_.Role -eq 'Measure' } | Select-Object -First 1)[0]
    $dimension = @($columns | Where-Object { $_.Role -eq 'Dimension' -and $_.DistinctCount -gt 1 -and $_.DistinctCount -le 50 } | Select-Object -First 1)[0]
    if (-not $dimension) {
        $dimension = @($columns | Where-Object { $_.Role -eq 'Time' } | Select-Object -First 1)[0]
    }

    $sourceLiteral = ConvertTo-ExcelAiPowerShellLiteral $SourcePath
    $workbookLiteral = ConvertTo-ExcelAiPowerShellLiteral $WorkbookPath
    $promptLiteral = ConvertTo-ExcelAiPowerShellLiteral $Prompt
    $sourceNameLiteral = ConvertTo-ExcelAiPowerShellLiteral $source.Name
    $measureLiteral = ConvertTo-ExcelAiPowerShellLiteral $(if ($measure) { $measure.Name } else { '' })
    $dimensionLiteral = ConvertTo-ExcelAiPowerShellLiteral $(if ($dimension) { $dimension.Name } else { '' })
    $pivotNameLiteral = ConvertTo-ExcelAiPowerShellLiteral (Get-ExcelAiSafeName -Name ("Pivot_$($source.Name)_$($dimension.Name)") -DefaultName 'PivotSummary')

    $formatLines = @()
    foreach ($column in $columns) {
        if ($column.SuggestedNumberFormat) {
            $formatLines += "        @{ Name = $(ConvertTo-ExcelAiPowerShellLiteral $column.Name); Format = $(ConvertTo-ExcelAiPowerShellLiteral $column.SuggestedNumberFormat) }"
        }
    }
    if ($formatLines.Count -eq 0) {
        $formatBlock = '    @()'
    }
    else {
        $formatBlock = "    @(`n" + ($formatLines -join ",`n") + "`n    )"
    }
    $chartDefaults = Get-ExcelAiPromptChartDefaults -Prompt $Prompt
    $chartTypeLiteral = ConvertTo-ExcelAiPowerShellLiteral $chartDefaults.ChartType
    $chartNoLegendLiteral = if ($chartDefaults.NoLegend) { '$true' } else { '$false' }
    $chartShowCategoryLiteral = if ($chartDefaults.ShowCategory) { '$true' } else { '$false' }
    $chartShowPercentLiteral = if ($chartDefaults.ShowPercent) { '$true' } else { '$false' }
    $conditionalFormatSpecs = @(Get-ExcelAiPromptConditionalFormatSpecs -Prompt $Prompt -DefaultToDataBar)
    $conditionalFormatLines = @()
    foreach ($formatSpec in $conditionalFormatSpecs) {
        $parts = @("Type = $(ConvertTo-ExcelAiPowerShellLiteral $formatSpec.Type)")
        foreach ($propertyName in 'Color', 'Rank', 'BackgroundColor', 'FontColor') {
            if ($null -ne $formatSpec.$propertyName) {
                if ($formatSpec.$propertyName -is [int]) {
                    $parts += "$propertyName = $($formatSpec.$propertyName)"
                }
                else {
                    $parts += "$propertyName = $(ConvertTo-ExcelAiPowerShellLiteral ([string]$formatSpec.$propertyName))"
                }
            }
        }
        if ($formatSpec.Bold) { $parts += 'Bold = $true' }
        $conditionalFormatLines += "        @{ $($parts -join '; ') }"
    }
    if ($conditionalFormatLines.Count -eq 0) {
        $conditionalFormatBlock = '    @()'
    }
    else {
        $conditionalFormatBlock = "    @(`n" + ($conditionalFormatLines -join ",`n") + "`n    )"
    }

    return @"
# Generated by ImportExcel AI.
# Prompt: $Prompt
[CmdletBinding()]
param(
    [string]`$SourcePath = $sourceLiteral,
    [string]`$OutputPath = $workbookLiteral,
    [string]`$ImportExcelModule = 'ImportExcel',
    [switch]`$Show,
    [switch]`$Force
)

`$ErrorActionPreference = 'Stop'
`$prompt = $promptLiteral
`$sourceName = $sourceNameLiteral
`$dimensionColumn = $dimensionLiteral
`$measureColumn = $measureLiteral
`$pivotName = $pivotNameLiteral
`$numberFormats =
$formatBlock
`$conditionalFormats =
$conditionalFormatBlock

if (-not (Get-Command Export-Excel -ErrorAction SilentlyContinue)) {
    if (Test-Path -LiteralPath `$ImportExcelModule) {
        Import-Module `$ImportExcelModule -Force
    }
    else {
        Import-Module ImportExcel -Force
    }
}

`$SourcePath = `$ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(`$SourcePath)
`$OutputPath = `$ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(`$OutputPath)

if ((Test-Path -LiteralPath `$OutputPath) -and -not `$Force) {
    throw "Output workbook already exists: `$OutputPath. Use -Force to overwrite it."
}
if (Test-Path -LiteralPath `$OutputPath) {
    Remove-Item -LiteralPath `$OutputPath -Force
}

`$extension = [System.IO.Path]::GetExtension(`$SourcePath).ToLowerInvariant()
switch (`$extension) {
    '.csv'  { `$data = @(Import-Csv -LiteralPath `$SourcePath) }
    '.tsv'  { `$data = @(Import-Csv -LiteralPath `$SourcePath -Delimiter "`t") }
    '.xlsx' { `$data = @(Import-Excel -Path `$SourcePath) }
    '.xlsm' { `$data = @(Import-Excel -Path `$SourcePath) }
    default { throw "Unsupported input file type: `$extension" }
}

if (`$data.Count -eq 0) {
    throw "No rows were found in `$SourcePath."
}

`$exportParams = @{
    Path = `$OutputPath
    WorksheetName = 'Data'
    TableName = 'SourceData'
    TableStyle = 'Medium6'
    AutoSize = `$true
    AutoFilter = `$true
    BoldTopRow = `$true
    FreezeTopRow = `$true
    PassThru = `$true
}
`$excel = `$data | Export-Excel @exportParams

try {
    `$ws = `$excel.Workbook.Worksheets['Data']
    `$ws.View.ShowGridLines = `$false

    `$headerMap = @{}
    for (`$column = 1; `$column -le `$ws.Dimension.End.Column; `$column++) {
        `$header = [string]`$ws.Cells[1, `$column].Value
        if (-not [string]::IsNullOrWhiteSpace(`$header)) {
            `$headerMap[`$header] = `$column
        }
    }

    foreach (`$format in `$numberFormats) {
        if (`$headerMap.ContainsKey(`$format.Name) -and `$ws.Dimension.End.Row -gt 1) {
            `$column = `$headerMap[`$format.Name]
            Set-ExcelRange -Address `$ws.Cells[2, `$column, `$ws.Dimension.End.Row, `$column] -NumberFormat `$format.Format
        }
    }

    if (`$measureColumn -and `$headerMap.ContainsKey(`$measureColumn) -and `$ws.Dimension.End.Row -gt 1) {
        `$measureIndex = `$headerMap[`$measureColumn]
        `$measureLetter = (Get-ExcelColumnName -ColumnNumber `$measureIndex).ColumnName
        `$measureAddress = `$measureLetter + '2:' + `$measureLetter + `$ws.Dimension.End.Row
        foreach (`$format in `$conditionalFormats) {
            switch ([string]`$format.Type) {
                'DataBar' {
                    `$color = if (`$format.Color) { [string]`$format.Color } else { 'SteelBlue' }
                    Add-ConditionalFormatting -Worksheet `$ws -Address `$measureAddress -DataBarColor `$color | Out-Null
                }
                'ThreeColorScale' {
                    Add-ConditionalFormatting -Worksheet `$ws -Address `$measureAddress -RuleType ThreeColorScale | Out-Null
                }
                'TwoColorScale' {
                    Add-ConditionalFormatting -Worksheet `$ws -Address `$measureAddress -RuleType TwoColorScale | Out-Null
                }
                'Top' {
                    `$rank = if (`$format.Rank) { [int]`$format.Rank } else { 10 }
                    `$topParams = @{
                        Worksheet = `$ws
                        Address = `$measureAddress
                        RuleType = 'Top'
                        ConditionValue = `$rank
                        BackgroundColor = if (`$format.BackgroundColor) { [string]`$format.BackgroundColor } else { 'LightGreen' }
                        ForegroundColor = if (`$format.FontColor) { [string]`$format.FontColor } else { 'DarkGreen' }
                    }
                    if (`$format.Bold) { `$topParams.Bold = `$true }
                    Add-ConditionalFormatting @topParams | Out-Null
                }
                'Bottom' {
                    `$rank = if (`$format.Rank) { [int]`$format.Rank } else { 10 }
                    `$bottomParams = @{
                        Worksheet = `$ws
                        Address = `$measureAddress
                        RuleType = 'Bottom'
                        ConditionValue = `$rank
                        BackgroundColor = if (`$format.BackgroundColor) { [string]`$format.BackgroundColor } else { 'LightPink' }
                        ForegroundColor = if (`$format.FontColor) { [string]`$format.FontColor } else { 'DarkRed' }
                    }
                    if (`$format.Bold) { `$bottomParams.Bold = `$true }
                    Add-ConditionalFormatting @bottomParams | Out-Null
                }
            }
        }
    }

    if (`$dimensionColumn -and `$measureColumn -and `$headerMap.ContainsKey(`$dimensionColumn) -and `$headerMap.ContainsKey(`$measureColumn)) {
        `$xColumn = `$headerMap[`$dimensionColumn]
        `$yColumn = `$headerMap[`$measureColumn]
        `$chartParams = @{
            Worksheet = `$ws
            Title = "`$measureColumn by `$dimensionColumn"
            ChartType = $chartTypeLiteral
            XRange = [OfficeOpenXml.ExcelAddress]::GetAddress(2, `$xColumn, `$ws.Dimension.End.Row, `$xColumn)
            YRange = [OfficeOpenXml.ExcelAddress]::GetAddress(2, `$yColumn, `$ws.Dimension.End.Row, `$yColumn)
            Width = 640
            Height = 360
            Row = 1
            Column = [Math]::Min(`$ws.Dimension.End.Column + 2, 12)
            NoLegend = $chartNoLegendLiteral
            ShowCategory = $chartShowCategoryLiteral
            ShowPercent = $chartShowPercentLiteral
        }
        Add-ExcelChart @chartParams | Out-Null

        `$pivotParams = @{
            ExcelPackage = `$excel
            PivotTableName = `$pivotName
            SourceWorkSheet = `$ws
            SourceRange = `$ws.Dimension.Address
            PivotRows = `$dimensionColumn
            PivotData = @{ `$measureColumn = 'Sum' }
            PivotTableStyle = 'Medium9'
            IncludePivotChart = `$true
            ChartType = $chartTypeLiteral
            ChartTitle = "`$measureColumn by `$dimensionColumn"
            NoLegend = $chartNoLegendLiteral
            ShowCategory = $chartShowCategoryLiteral
            ShowPercent = $chartShowPercentLiteral
        }
        Add-PivotTable @pivotParams | Out-Null
    }

    `$summary = Add-Worksheet -ExcelPackage `$excel -WorksheetName 'Summary' -MoveToEnd
    `$summary.View.ShowGridLines = `$false
    `$summary.Cells['A1'].Value = 'Generated Report'
    `$summary.Cells['A2'].Value = 'Prompt'
    `$summary.Cells['B2'].Value = `$prompt
    `$summary.Cells['A3'].Value = 'Source'
    `$summary.Cells['B3'].Value = `$SourcePath
    `$summary.Cells['A4'].Value = 'Rows'
    `$summary.Cells['B4'].Value = `$data.Count
    `$summary.Cells['A5'].Value = 'Columns'
    `$summary.Cells['B5'].Value = `$ws.Dimension.End.Column
    Set-ExcelRange -Address `$summary.Cells['A1:B1'] -Bold -FontSize 16
    Set-ExcelRange -Address `$summary.Cells['A2:A5'] -Bold
    `$summary.Cells.AutoFitColumns()
}
finally {
    Close-ExcelPackage -ExcelPackage `$excel -Show:`$Show
}

[pscustomobject]@{
    Path = `$OutputPath
    SourcePath = `$SourcePath
    ScriptPath = `$PSCommandPath
}
"@
}

function Get-ExcelAiColumnIndex {
    param(
        [Parameter(Mandatory)]
        [OfficeOpenXml.ExcelWorksheet]$Worksheet,

        [Parameter(Mandatory)]
        [string]$HeaderName,

        [int]$HeaderRow = 1
    )

    if (-not $Worksheet.Dimension) { return 0 }

    for ($column = $Worksheet.Dimension.Start.Column; $column -le $Worksheet.Dimension.End.Column; $column++) {
        if ([string]$Worksheet.Cells[$HeaderRow, $column].Value -eq $HeaderName) {
            return $column
        }
    }

    return 0
}
