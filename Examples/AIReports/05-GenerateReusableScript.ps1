param(
    [string]$Model = $(if ($env:IMPORTEXCEL_AI_MODEL) { $env:IMPORTEXCEL_AI_MODEL } elseif ($env:PSAISUITE_DEFAULT_MODEL) { $env:PSAISUITE_DEFAULT_MODEL } else { 'openai:gpt-5.2' }),
    [switch]$NoAI,
    [switch]$Show,
    [switch]$ShowFallbackWarnings
)

try { Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force } catch { throw ; return }

$exampleRoot = Join-Path $env:TEMP 'ImportExcelAIExamples'
$null = New-Item -Path $exampleRoot -ItemType Directory -Force

$csvPath = Join-Path $exampleRoot 'agent-sales.csv'
$reportPath = Join-Path $exampleRoot 'agent-sales-report.xlsx'
$scriptPath = Join-Path $exampleRoot 'build-agent-sales-report.ps1'

@"
Region,State,Product,Units,Revenue,Margin,CloseDate
West,Texas,Widget,927,92371,0.28,2024-01-15
North,Tennessee,Gadget,466,77067,0.21,2024-01-18
East,Florida,Widget,520,45868,0.18,2024-02-05
East,Maine,Service,828,66124,0.32,2024-02-09
West,Virginia,Gadget,465,5358,0.08,2024-03-11
North,Missouri,Widget,436,23567,0.12,2024-03-21
South,Kansas,Service,214,99247,0.34,2024-04-04
North,North Dakota,Gadget,789,64072,0.25,2024-04-15
South,Delaware,Widget,712,50855,0.19,2024-05-01
"@ | Set-Content -Path $csvPath

$agentParams = @{
    Path = $csvPath
    # Prompt = 'Create a polished executive sales workbook. Include a source data table, chart, pivot analysis, summary notes, and clean formatting. Keep the generated PowerShell script for future runs.'
    Prompt = 'Create a polished executive sales workbook. Include a source data table, pie chartKeep the generated PowerShell script for future runs.'
    OutputPath = $reportPath
    ScriptPath = $scriptPath
    Model = $Model
    FallbackToDefault = $true
    NoAI = $NoAI
    Force = $true
    Show = $Show
}
if (-not $ShowFallbackWarnings) {
    $agentParams.WarningAction = 'SilentlyContinue'
}

$previousWarningPreference = $WarningPreference
if (-not $ShowFallbackWarnings) {
    $WarningPreference = 'SilentlyContinue'
}
try {
    $result = Invoke-ExcelAgent @agentParams
}
finally {
    $WarningPreference = $previousWarningPreference
}

[pscustomobject]@{
    ReportPath = $result.Path
    ScriptPath = $result.ScriptPath
    UsedAI = $result.UsedAI
    RerunExample = "& '$($result.ScriptPath)' -SourcePath '$csvPath' -OutputPath '$reportPath' -Force"
}
