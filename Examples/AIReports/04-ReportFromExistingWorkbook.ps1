param(
    [switch]$Show
)

try { Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force } catch { throw ; return }

$exampleRoot = Join-Path $env:TEMP 'ImportExcelAIExamples'
$null = New-Item -Path $exampleRoot -ItemType Directory -Force

$sourceWorkbook = Join-Path $exampleRoot 'department-source.xlsx'
$reportPath = Join-Path $exampleRoot 'department-workbook-report.xlsx'
Remove-Item -Path $sourceWorkbook, $reportPath -ErrorAction Ignore

$sales = @"
Department,Owner,Quarter,Bookings,Revenue,Forecast
Sales,Ada,Q1,24,182000,210000
Sales,Ada,Q2,31,244000,260000
Services,Grace,Q1,18,138000,150000
Services,Grace,Q2,22,169000,190000
Training,Katherine,Q1,11,72000,80000
Training,Katherine,Q2,15,96000,110000
"@ | ConvertFrom-Csv

$expenses = @"
Department,Quarter,Travel,Software,Contractors
Sales,Q1,23000,8500,42000
Sales,Q2,26000,9200,41000
Services,Q1,18000,12000,55000
Services,Q2,21000,13500,62000
Training,Q1,9000,6100,18000
Training,Q2,11000,7200,22000
"@ | ConvertFrom-Csv

$sales | Export-Excel -Path $sourceWorkbook -WorksheetName Sales -AutoSize -TableName SalesData -TableStyle Medium6
$expenses | Export-Excel -Path $sourceWorkbook -WorksheetName Expenses -AutoSize -TableName ExpenseData -TableStyle Medium4

$result = Invoke-ExcelPrompt `
    -Path $sourceWorkbook `
    -Prompt 'Create a management report from the workbook. Summarize the available sheets and create useful formatted data tabs, charts, and pivots.' `
    -OutputPath $reportPath `
    -NoAI `
    -Force `
    -Show:$Show

[pscustomobject]@{
    SourceWorkbook = $sourceWorkbook
    ReportPath = $result.Path
    Worksheets = $result.Worksheets -join ', '
}

