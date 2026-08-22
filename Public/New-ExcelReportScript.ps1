function New-ExcelReportScript {
    <#
    .SYNOPSIS
        Generates a reusable PowerShell script for an AI-assisted Excel report.

    .DESCRIPTION
        Profiles the input data, then asks PSAISuite to generate ImportExcel-based
        PowerShell automation for the requested report. Use -NoAI to generate a local
        deterministic starter script. The generated script is parse-checked and checked
        for supported commands before it is written to disk.

    .PARAMETER Path
        Source CSV/TSV/workbook path.

    .PARAMETER Prompt
        Natural-language request for the generated report script.

    .PARAMETER OutputPath
        Workbook path that the generated script should use by default.

    .PARAMETER ScriptPath
        Optional path where the generated .ps1 should be saved.

    .PARAMETER NoAI
        Generates a deterministic local script instead of calling PSAISuite.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Prompt,

        [string]$OutputPath,

        [string]$ScriptPath,

        [string[]]$WorksheetName,

        [char]$Delimiter = ',',

        [string]$Model = $(if ($env:IMPORTEXCEL_AI_MODEL) { $env:IMPORTEXCEL_AI_MODEL } elseif ($env:PSAISUITE_DEFAULT_MODEL) { $env:PSAISUITE_DEFAULT_MODEL } else { 'openai:gpt-4o-mini' }),

        [string]$PSAISuitePath,

        [switch]$NoAI,

        [switch]$FallbackToDefault,

        [switch]$Force
    )

    process {
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        if (-not $OutputPath) {
            $directory = Split-Path -Path $resolvedPath -Parent
            $fileName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath) + '-agent-report.xlsx'
            $OutputPath = Join-Path -Path $directory -ChildPath $fileName
        }
        $resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)

        $summary = Get-ExcelDatasetSummary -Path $resolvedPath -WorksheetName $WorksheetName -Delimiter $Delimiter
        $defaultScript = New-ExcelAiDefaultReportScript -DatasetSummary $summary -SourcePath $resolvedPath -WorkbookPath $resolvedOutputPath -Prompt $Prompt
        $usedAI = $false

        if ($NoAI) {
            $script = $defaultScript
        }
        else {
            $summaryJson = $summary | ConvertTo-Json -Depth 20
            $systemPrompt = @'
You generate reusable PowerShell scripts that automate Excel reports with ImportExcel.
Return only a complete PowerShell script. Do not use markdown. Do not explain.
The script must:
- start with a param block exposing SourcePath, OutputPath, Show, and Force
- import ImportExcel when needed
- never use Excel COM automation
- never install modules
- never delete anything except OutputPath when Force is supplied
- create an .xlsx workbook using ImportExcel commands
- save and return an object with Path, SourcePath, and ScriptPath when it finishes
Use these ImportExcel patterns:
- create/open the workbook with `$excel = $data | Export-Excel -Path $OutputPath -WorksheetName 'Data' -PassThru`
- apply table filters with `Export-Excel -AutoFilter`; do not use `Set-ExcelRange -AutoFilter`
- get worksheets with `$excel.Workbook.Worksheets['Data']`
- add worksheets with `Add-Worksheet -ExcelPackage $excel -WorksheetName 'Summary'` or `$excel.Workbook.Worksheets.Add('Summary')`
- format cells with `Set-ExcelRange -Address $worksheet.Cells['A1:B2'] ...` or `Set-ExcelRange -Worksheet $worksheet -Range 'A1:B2' ...`
- use `-BackgroundColor` for fills and explicit border parameters with values, such as `-BorderAround Thin`, `-BorderBottom Thick`, `-BorderTop Thin`, `-BorderLeft Thin`, or `-BorderRight Thin`
- use ImportExcel color parameters like `-BackgroundColor`, `-FontColor`, and `-BorderColor`; do not call EPPlus `.SetColor(...)` methods directly
- add charts with `Add-ExcelChart -Worksheet $worksheet -ChartType <requested chart type> -XRange ([OfficeOpenXml.ExcelAddress]::GetAddress(2, $xColumn, $worksheet.Dimension.End.Row, $xColumn)) -YRange ([OfficeOpenXml.ExcelAddress]::GetAddress(2, $yColumn, $worksheet.Dimension.End.Row, $yColumn))`; use address strings for XRange and YRange, not ExcelRange cell objects
- honor explicit chart requests in the user prompt, such as Pie, Doughnut, Line, BarClustered, or ColumnClustered
- honor formatting requests in the user prompt with ImportExcel APIs: data bars use `Add-ConditionalFormatting -Worksheet $worksheet -Address $range -DataBarColor SteelBlue`; color scales use `-RuleType ThreeColorScale`; top/bottom highlighting uses `-RuleType Top` or `-RuleType Bottom` with `-ConditionValue`
- add pivots with `Add-PivotTable -ExcelPackage $excel -PivotTableName 'PivotName' -SourceWorksheet $sourceWorksheet -SourceRange $sourceWorksheet.Dimension.Address -PivotRows 'Category' -PivotData @{ Amount = 'Sum' }`
- close/save with `Close-ExcelPackage -ExcelPackage $excel -Show:$Show`
Do not use these non-ImportExcel or unsupported shapes:
- New-ExcelPackage or Save-ExcelPackage
- Add-Worksheet -Workbook
- Set-ExcelRange -ExcelPackage or Set-ExcelRange -WorksheetName
- Set-ExcelRange -FillColor, Set-ExcelRange -Border, or bare border parameters without a style value such as `-BorderAround`
- Set-ExcelRange -AutoFilter
- direct EPPlus styling calls like `.Style.Font.Color.SetColor(...)` or `.Style.Fill.BackgroundColor.SetColor(...)`
- Add-PivotTable -DestinationWorksheet
- New-ExcelChartDefinition -PivotTable
- Add-ExcelChart -Chart
Prefer expressive ImportExcel automation over a fixed JSON plan. Use pivots, charts, formulas, helper sheets, and formatting when useful.
'@

            $userPrompt = @"
Prompt:
$Prompt

Default source path:
$resolvedPath

Default workbook output path:
$resolvedOutputPath

Dataset summary:
$summaryJson
"@

            try {
                Import-ExcelAiPSAISuite -PSAISuitePath $PSAISuitePath
                $messages = @(
                    @{ role = 'system'; content = $systemPrompt },
                    @{ role = 'user'; content = $userPrompt }
                )
                $response = Invoke-ChatCompletion -Messages $messages -Model $Model
                $script = ConvertFrom-ExcelAiPowerShellResponse -Text $response
                $usedAI = $true
            }
            catch {
                if ($FallbackToDefault) {
                    Write-Warning "Falling back to the deterministic report script because AI script generation failed: $_"
                    $script = $defaultScript
                    $usedAI = $false
                }
                else {
                    throw
                }
            }
        }

        $parseErrors = @(Test-ExcelAiPowerShellScript -Script $script)
        if ($parseErrors.Count -gt 0) {
            if ($FallbackToDefault -and -not $NoAI) {
                Write-Warning "Falling back to the deterministic report script because the generated script did not parse."
                $script = $defaultScript
                $usedAI = $false
                $parseErrors = @(Test-ExcelAiPowerShellScript -Script $script)
            }
            if ($parseErrors.Count -gt 0) {
                $message = ($parseErrors | Select-Object -First 3 | ForEach-Object { $_.Message }) -join '; '
                throw "Generated script failed PowerShell parsing: $message"
            }
        }

        $usageIssues = @(Test-ExcelAiPowerShellCommandUsage -Script $script)
        if ($usageIssues.Count -gt 0) {
            if ($FallbackToDefault -and -not $NoAI) {
                $issueText = ($usageIssues | Select-Object -First 3 | ForEach-Object { $_.Message }) -join '; '
                Write-Warning "Falling back to the deterministic report script because the generated script used unsupported commands or parameters: $issueText"
                $script = $defaultScript
                $usedAI = $false
                $usageIssues = @(Test-ExcelAiPowerShellCommandUsage -Script $script)
            }
            if ($usageIssues.Count -gt 0) {
                $message = ($usageIssues | Select-Object -First 5 | ForEach-Object { $_.Message }) -join '; '
                throw "Generated script used unsupported commands or parameters: $message"
            }
        }

        if ($ScriptPath) {
            $resolvedScriptPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ScriptPath)
            if ((Test-Path -LiteralPath $resolvedScriptPath) -and -not $Force) {
                throw "Script already exists: $resolvedScriptPath. Use -Force to overwrite it."
            }
            $scriptDirectory = Split-Path -Path $resolvedScriptPath -Parent
            if ($scriptDirectory -and -not (Test-Path -LiteralPath $scriptDirectory)) {
                $null = New-Item -Path $scriptDirectory -ItemType Directory -Force
            }
            Set-Content -LiteralPath $resolvedScriptPath -Value $script -Encoding UTF8
        }
        else {
            $resolvedScriptPath = $null
        }

        [pscustomobject][ordered]@{
            ScriptPath = $resolvedScriptPath
            WorkbookPath = $resolvedOutputPath
            SourcePath = $resolvedPath
            Model = if ($usedAI) { $Model } else { $null }
            UsedAI = $usedAI
            Script = $script
            DatasetSummary = $summary
        }
    }
}
