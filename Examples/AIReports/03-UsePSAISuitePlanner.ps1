param(
    [string]$Model = $(if ($env:IMPORTEXCEL_AI_MODEL) { $env:IMPORTEXCEL_AI_MODEL } elseif ($env:PSAISUITE_DEFAULT_MODEL) { $env:PSAISUITE_DEFAULT_MODEL } else { 'openai:gpt-5.4-mini' }),
    [string]$PSAISuitePath = 'D:\mygit\psaisuite\PSAISuite.psd1',
    [switch]$Show
)

try { Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force } catch { throw ; return }

$exampleRoot = Join-Path $env:TEMP 'ImportExcelAIExamples'
$null = New-Item -Path $exampleRoot -ItemType Directory -Force

$csvPath = Join-Path $exampleRoot 'support-tickets.csv'
$reportPath = Join-Path $exampleRoot 'support-tickets-ai-report.xlsx'

@"
Queue,Priority,Product,Opened,Closed,HoursToResolve,Satisfaction
Enterprise,High,API,2024-01-02,2024-01-03,21,4.7
Enterprise,Critical,Connectors,2024-01-08,2024-01-09,16,4.9
Commercial,Medium,Desktop,2024-01-13,2024-01-15,37,4.3
Commercial,Low,API,2024-01-21,2024-01-23,52,4.1
Community,Medium,Import,2024-02-04,2024-02-05,18,4.5
Enterprise,High,Desktop,2024-02-14,2024-02-15,25,4.6
Commercial,High,Connectors,2024-03-01,2024-03-02,20,4.4
Community,Low,Import,2024-03-12,2024-03-15,72,3.9
Enterprise,Critical,API,2024-04-07,2024-04-08,13,4.8
"@ | Set-Content -Path $csvPath

$invokeParams = @{
    Path = $csvPath
    Prompt = 'Create an operations dashboard for support leaders. Prioritize SLA risk, resolution time, customer satisfaction, and queue/product trends.'
    OutputPath = $reportPath
    Model = $Model
    FallbackToDefault = $true
    Force = $true
    Show = $Show
}

if (Test-Path -LiteralPath $PSAISuitePath) {
    $invokeParams.PSAISuitePath = $PSAISuitePath
}

$result = Invoke-ExcelPrompt @invokeParams

[pscustomobject]@{
    Model = $Model
    ReportPath = $result.Path
    Worksheets = $result.Worksheets -join ', '
}

