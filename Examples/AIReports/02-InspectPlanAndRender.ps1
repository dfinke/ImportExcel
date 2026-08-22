param(
    [switch]$Show
)

try { Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force } catch { throw ; return }

$exampleRoot = Join-Path $env:TEMP 'ImportExcelAIExamples'
$null = New-Item -Path $exampleRoot -ItemType Directory -Force

$csvPath = Join-Path $exampleRoot 'inventory-health.csv'
$planPath = Join-Path $exampleRoot 'inventory-health-plan.json'
$reportPath = Join-Path $exampleRoot 'inventory-health-report.xlsx'

@"
Warehouse,Category,Sku,OnHand,ReorderPoint,UnitCost,LastReceipt
Northeast,Hardware,HW-100,120,50,18.25,2024-01-04
Northeast,Hardware,HW-200,18,40,31.50,2024-01-17
Northeast,Accessories,AC-100,240,100,7.99,2024-02-12
Central,Hardware,HW-100,75,50,18.25,2024-02-19
Central,Accessories,AC-200,12,75,11.50,2024-03-03
West,Hardware,HW-300,42,25,99.00,2024-03-15
West,Accessories,AC-100,310,100,7.99,2024-04-07
South,Hardware,HW-200,22,40,31.50,2024-04-21
South,Accessories,AC-300,95,80,13.75,2024-05-02
"@ | Set-Content -Path $csvPath

$summary = Get-ExcelDatasetSummary -Path $csvPath
$plan = New-ExcelReportPlan `
    -DatasetSummary $summary `
    -Prompt 'Create an inventory health report. Highlight low stock risks and summarize inventory value by warehouse.' `
    -NoAI

$plan | ConvertTo-Json -Depth 20 | Set-Content -Path $planPath

$result = Invoke-ExcelReportPlan `
    -Plan $plan `
    -Path $reportPath `
    -SourcePath $csvPath `
    -DatasetSummary $summary `
    -Force `
    -Show:$Show

[pscustomobject]@{
    DatasetSummary = $summary
    PlanPath = $planPath
    ReportPath = $result.Path
    Worksheets = $result.Worksheets -join ', '
}

