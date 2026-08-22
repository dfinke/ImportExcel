function New-ExcelReportPlan {
    <#
    .SYNOPSIS
        Creates a constrained ImportExcel report plan from a dataset summary and prompt.

    .DESCRIPTION
        Builds a JSON-compatible plan that Invoke-ExcelReportPlan can execute. By default
        the function uses PSAISuite's Invoke-ChatCompletion to translate the prompt and
        dataset summary into a plan. Use -NoAI to create a deterministic starter plan
        without calling a model.

    .PARAMETER DatasetSummary
        Output from Get-ExcelDatasetSummary.

    .PARAMETER Prompt
        Natural-language instructions for the report.

    .PARAMETER Model
        PSAISuite model identifier, such as openai:gpt-4o-mini or anthropic:claude-sonnet-4-6.

    .PARAMETER PSAISuitePath
        Optional path to PSAISuite.psd1. If omitted, the function tries an already-loaded
        PSAISuite module, $env:PSAISUITE_PATH, a sibling psaisuite repo, then PSModulePath.

    .PARAMETER NoAI
        Skips PSAISuite and returns a deterministic starter plan.

    .PARAMETER FallbackToDefault
        If the model call fails or returns invalid JSON, returns the deterministic starter plan.

    .PARAMETER AsJson
        Returns the plan as JSON.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $DatasetSummary,

        [Parameter(Mandatory)]
        [string]$Prompt,

        [string]$Model = $(if ($env:IMPORTEXCEL_AI_MODEL) { $env:IMPORTEXCEL_AI_MODEL } elseif ($env:PSAISUITE_DEFAULT_MODEL) { $env:PSAISUITE_DEFAULT_MODEL } else { 'openai:gpt-4o-mini' }),

        [string]$PSAISuitePath,

        [switch]$NoAI,

        [switch]$FallbackToDefault,

        [switch]$AsJson,

        [switch]$Compress
    )

    process {
        $defaultPlan = Get-ExcelAiDefaultReportPlan -DatasetSummary $DatasetSummary -Prompt $Prompt

        if ($NoAI) {
            if ($AsJson) { return $defaultPlan | ConvertTo-Json -Depth 20 -Compress:$Compress }
            return $defaultPlan
        }

        $summaryJson = $DatasetSummary | ConvertTo-Json -Depth 20
        $systemPrompt = @'
You create executable ImportExcel report plans.
Return JSON only. Do not use markdown. Do not explain.
Use only column names that appear in the dataset summary.
The JSON shape must be:
{
  "Version": "1.0",
  "WorkbookTitle": "short title",
  "SourcePath": "input path",
  "Summary": { "SheetName": "Summary", "Title": "title", "IncludeDatasetProfile": true },
  "Tables": [
    {
      "SourceName": "source name from summary",
      "SheetName": "worksheet name",
      "TableName": "letters_numbers_underscore",
      "TableStyle": "Medium6",
      "AutoSize": true,
      "AutoFilter": true,
      "BoldTopRow": true,
      "FreezeTopRow": true,
      "NumberFormats": { "ColumnName": "Currency" }
    }
  ],
  "Charts": [
    {
      "SourceName": "source name",
      "Title": "chart title",
      "ChartType": "ColumnClustered",
      "XColumn": "dimension column",
      "YColumn": ["measure column"],
      "Width": 640,
      "Height": 360,
      "Row": 1,
      "Column": 7,
      "NoLegend": true
    }
  ],
  "Pivots": [
    {
      "SourceName": "source name",
      "PivotTableName": "letters_numbers_underscore",
      "PivotRows": ["dimension column"],
      "PivotColumns": [],
      "PivotData": { "measure column": "Sum" },
      "PivotTableStyle": "Medium9",
      "IncludePivotChart": true,
      "ChartType": "ColumnClustered",
      "ChartTitle": "pivot chart title",
      "NoLegend": true
    }
  ],
  "ConditionalFormats": [
    { "SourceName": "source name", "Column": "measure column", "Type": "DataBar", "Color": "SteelBlue" }
  ],
  "AnalysisSheets": [
    { "SourceName": "source name", "Type": "ExecutiveDashboard", "SheetName": "Executive Dashboard", "Title": "Executive Dashboard", "TableStyle": "Medium4" },
    { "SourceName": "source name", "Type": "DataScience", "SheetName": "Statistical Analysis", "Title": "Statistical Analysis", "TableStyle": "Medium7" },
    { "SourceName": "source name", "Type": "CorrelationMatrix", "SheetName": "Correlation Matrix", "Title": "Correlation Matrix", "TableStyle": "Light11" }
  ]
}
Valid pivot functions are Sum, Count, Average, Min, and Max.
Use at most one useful chart and one useful pivot per source unless the prompt clearly asks for more.
For executive/dashboard prompts, include an ExecutiveDashboard analysis sheet.
For data scientist/statistical prompts, include DataScience and CorrelationMatrix analysis sheets when numeric columns exist.
'@

        $userPrompt = @"
Prompt:
$Prompt

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
            $plan = ConvertFrom-ExcelAiJsonResponse -Text $response
        }
        catch {
            if ($FallbackToDefault) {
                Write-Warning "Falling back to the deterministic report plan because AI planning failed: $_"
                $plan = $defaultPlan
            }
            else {
                throw
            }
        }

        $plan = Add-ExcelAiPromptAnalysisDefaults -Plan $plan -DatasetSummary $DatasetSummary -Prompt $Prompt

        if ($AsJson) { return $plan | ConvertTo-Json -Depth 20 -Compress:$Compress }
        return $plan
    }
}
