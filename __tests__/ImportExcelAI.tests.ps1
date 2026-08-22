Import-Module $PSScriptRoot\..\ImportExcel.psd1 -Force

Describe "ImportExcel AI report planning" -Tag ImportExcelAI {
    BeforeAll {
        . (Join-Path $PSScriptRoot '..\Private\ImportExcelAI.Helpers.ps1')

        $script:csvPath = "TestDrive:\sales.csv"
        @"
Region,State,Units,Revenue,CloseDate
West,Texas,927,923.71,2024-01-15
North,Tennessee,466,770.67,2024-01-18
East,Florida,520,458.68,2024-02-05
East,Maine,828,661.24,2024-02-09
West,Virginia,465,53.58,2024-03-11
North,Missouri,436,235.67,2024-03-21
South,Kansas,214,992.47,2024-04-04
North,North Dakota,789,640.72,2024-04-15
South,Delaware,712,508.55,2024-05-01
"@ | Set-Content -Path $script:csvPath
    }

    It "Profiles a CSV file for report planning" {
        $summary = Get-ExcelDatasetSummary -Path $script:csvPath

        $summary.FileName | Should -BeExactly "sales.csv"
        $summary.Sources.Count | Should -Be 1
        $summary.Sources[0].RowCount | Should -Be 9
        $summary.Sources[0].ColumnCount | Should -Be 5
        $summary.Sources[0].ColumnSummaries.Name | Should -Contain "Revenue"
        ($summary.Sources[0].ColumnSummaries | Where-Object Name -eq "Revenue").Role | Should -BeExactly "Measure"
        ($summary.Sources[0].ColumnSummaries | Where-Object Name -eq "CloseDate").Role | Should -BeExactly "Time"
    }

    It "Creates a deterministic report plan without calling an AI model" {
        $summary = Get-ExcelDatasetSummary -Path $script:csvPath
        $plan = New-ExcelReportPlan -DatasetSummary $summary -Prompt "Create an executive sales report." -NoAI

        $plan.Version | Should -BeExactly "1.0"
        $plan.Tables.Count | Should -Be 1
        $plan.Tables[0].SourceName | Should -BeExactly "sales"
        $plan.Charts.Count | Should -BeGreaterThan 0
        $plan.Pivots.Count | Should -BeGreaterThan 0
    }

    It "Executes a report plan into a workbook" {
        $summary = Get-ExcelDatasetSummary -Path $script:csvPath
        $plan = New-ExcelReportPlan -DatasetSummary $summary -Prompt "Create an executive sales report." -NoAI
        $outputPath = "TestDrive:\sales-report.xlsx"

        $result = Invoke-ExcelReportPlan -Plan $plan -Path $outputPath -SourcePath $script:csvPath -DatasetSummary $summary -Force

        Test-Path -Path $outputPath | Should -Be $true
        $result.Path | Should -BeLike "*sales-report.xlsx"

        $excel = Open-ExcelPackage -Path $outputPath
        try {
            $excel.Workbook.Worksheets["Summary"] | Should -Not -BeNullOrEmpty
            $excel.Workbook.Worksheets["Column Profile"] | Should -Not -BeNullOrEmpty
            $excel.Workbook.Worksheets["sales"] | Should -Not -BeNullOrEmpty
            $excel.Workbook.Worksheets["sales"].Tables.Count | Should -Be 1
            $excel.Workbook.Worksheets[1].Name | Should -BeExactly "sales"
            $excel.Workbook.Worksheets[2].Name | Should -BeLike "Pivot*"
            $excel.Workbook.Worksheets[$excel.Workbook.Worksheets.Count - 1].Name | Should -BeExactly "Summary"
            $excel.Workbook.Worksheets[$excel.Workbook.Worksheets.Count].Name | Should -BeExactly "Column Profile"
        }
        finally {
            Close-ExcelPackage -ExcelPackage $excel -NoSave
        }

        $expandedPath = "TestDrive:\sales-report-expanded"
        Expand-Archive -Path $outputPath -DestinationPath $expandedPath -Force
        $pivotCache = Get-Content -Path (Join-Path $expandedPath 'xl\pivotCache\pivotCacheDefinition1.xml') -Raw
        $pivotCache | Should -Not -Match 'ref=""'
        $pivotCache | Should -Match 'worksheetSource'
    }

    It "Runs the end-to-end prompt workflow without AI when requested" {
        $outputPath = "TestDrive:\sales-prompt-report.xlsx"

        $result = Invoke-ExcelPrompt -Path $script:csvPath -Prompt "Create a sales report." -OutputPath $outputPath -NoAI -Force

        Test-Path -Path $outputPath | Should -Be $true
        $result.Worksheets | Should -Contain "Summary"
        $result.Worksheets | Should -Contain "sales"
    }

    It "Creates prompt-specific analysis sheets" {
        $summary = Get-ExcelDatasetSummary -Path $script:csvPath

        $executivePlan = New-ExcelReportPlan -DatasetSummary $summary -Prompt "Create an executive dashboard." -NoAI
        $executivePlan.AnalysisSheets.Type | Should -Contain "ExecutiveDashboard"

        $sciencePlan = New-ExcelReportPlan -DatasetSummary $summary -Prompt "Analyze the data as a data scientist." -NoAI
        $sciencePlan.AnalysisSheets.Type | Should -Contain "DataScience"
        $sciencePlan.AnalysisSheets.Type | Should -Contain "CorrelationMatrix"

        $outputPath = "TestDrive:\sales-science-report.xlsx"
        $result = Invoke-ExcelPrompt -Path $script:csvPath -Prompt "Analyze the data as a data scientist." -OutputPath $outputPath -NoAI -Force

        $result.Worksheets | Should -Contain "Statistical Analysis"
        $result.Worksheets | Should -Contain "Correlation Matrix"
    }

    It "Generates a reusable PowerShell report script" {
        $scriptPath = "TestDrive:\build-sales-report.ps1"
        $workbookPath = "TestDrive:\agent-script-report.xlsx"

        $result = New-ExcelReportScript -Path $script:csvPath -Prompt "Create a reusable sales report script." -OutputPath $workbookPath -ScriptPath $scriptPath -NoAI -Force

        Test-Path -Path $scriptPath | Should -Be $true
        $result.Script | Should -Match "Generated by ImportExcel AI"
        $result.Script | Should -Match "Export-Excel"

        $tokens = $null
        $parseErrors = $null
        [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $scriptPath).ProviderPath, [ref]$tokens, [ref]$parseErrors) | Out-Null
        @($parseErrors).Count | Should -Be 0
    }

    It "Runs the Excel agent and keeps the generated script" {
        $scriptPath = "TestDrive:\agent-build-sales-report.ps1"
        $workbookPath = "TestDrive:\agent-sales-report.xlsx"

        $result = Invoke-ExcelAgent -Path $script:csvPath -Prompt "Build an executive sales workbook and keep the script." -OutputPath $workbookPath -ScriptPath $scriptPath -NoAI -Force

        Test-Path -Path $scriptPath | Should -Be $true
        Test-Path -Path $workbookPath | Should -Be $true
        $result.ScriptPath | Should -BeLike "*agent-build-sales-report.ps1"
        $result.Path | Should -BeLike "*agent-sales-report.xlsx"

        $excel = Open-ExcelPackage -Path $workbookPath
        try {
            $excel.Workbook.Worksheets["Data"] | Should -Not -BeNullOrEmpty
            $excel.Workbook.Worksheets["Summary"] | Should -Not -BeNullOrEmpty
        }
        finally {
            Close-ExcelPackage -ExcelPackage $excel -NoSave
        }

        $expandedPath = "TestDrive:\agent-sales-report-expanded"
        Expand-Archive -Path $workbookPath -DestinationPath $expandedPath -Force
        $chartXml = Get-Content -Path (Join-Path $expandedPath 'xl\charts\chart1.xml') -Raw
        $chartXml | Should -Not -Match '#REF!'
        $chartXml | Should -Match '''Data''!\$?A\$?2:\$?A\$?10'
        $chartXml | Should -Match '''Data''!\$?C\$?2:\$?C\$?10'
    }

    It "Honors requested chart type when falling back to deterministic scripts" {
        $prompt = "Create a polished executive sales workbook. Include a source data table, pie chartKeep the generated PowerShell script for future runs."
        $scriptPath = "TestDrive:\agent-build-pie-report.ps1"
        $workbookPath = "TestDrive:\agent-pie-report.xlsx"

        $summary = Get-ExcelDatasetSummary -Path $script:csvPath
        $plan = New-ExcelReportPlan -DatasetSummary $summary -Prompt $prompt -NoAI
        $plan.Charts[0].ChartType | Should -BeExactly 'Pie'
        $plan.Pivots[0].ChartType | Should -BeExactly 'Pie'

        $scriptResult = New-ExcelReportScript -Path $script:csvPath -Prompt $prompt -OutputPath $workbookPath -ScriptPath $scriptPath -NoAI -Force
        $scriptResult.Script | Should -Match "ChartType = 'Pie'"

        Invoke-ExcelAgent -Path $script:csvPath -Prompt $prompt -OutputPath $workbookPath -ScriptPath $scriptPath -NoAI -Force | Out-Null

        $expandedPath = "TestDrive:\agent-pie-report-expanded"
        Expand-Archive -Path $workbookPath -DestinationPath $expandedPath -Force
        $chartXml = Get-Content -Path (Join-Path $expandedPath 'xl\charts\chart1.xml') -Raw
        $chartXml | Should -Match 'pieChart'
        $chartXml | Should -Not -Match '#REF!'
    }

    It "Honors prompt-requested conditional formatting in deterministic plans and scripts" {
        $prompt = "Create the data in a table, use data bars, add a color scale, highlight top 3 and bottom 2 values."
        $scriptPath = "TestDrive:\agent-build-format-report.ps1"
        $workbookPath = "TestDrive:\agent-format-report.xlsx"

        $summary = Get-ExcelDatasetSummary -Path $script:csvPath
        $plan = New-ExcelReportPlan -DatasetSummary $summary -Prompt $prompt -NoAI
        $plan.Tables[0].AutoFilter | Should -Be $true
        $plan.Tables[0].FreezeTopRow | Should -Be $true
        $plan.ConditionalFormats.Type | Should -Contain 'DataBar'
        $plan.ConditionalFormats.Type | Should -Contain 'ThreeColorScale'
        $plan.ConditionalFormats.Type | Should -Contain 'Top'
        $plan.ConditionalFormats.Type | Should -Contain 'Bottom'
        ($plan.ConditionalFormats | Where-Object Type -eq 'Top').Rank | Should -Be 3
        ($plan.ConditionalFormats | Where-Object Type -eq 'Bottom').Rank | Should -Be 2

        $scriptResult = New-ExcelReportScript -Path $script:csvPath -Prompt $prompt -OutputPath $workbookPath -ScriptPath $scriptPath -NoAI -Force
        $scriptResult.Script | Should -Match 'DataBarColor'
        $scriptResult.Script | Should -Match 'ThreeColorScale'
        $scriptResult.Script | Should -Match "RuleType = 'Top'"
        $scriptResult.Script | Should -Match "RuleType = 'Bottom'"

        Invoke-ExcelAgent -Path $script:csvPath -Prompt $prompt -OutputPath $workbookPath -ScriptPath $scriptPath -NoAI -Force | Out-Null

        $excel = Open-ExcelPackage -Path $workbookPath
        try {
            $worksheet = $excel.Workbook.Worksheets['Data']
            @($worksheet.ConditionalFormatting).Count | Should -BeGreaterOrEqual 4
            @($worksheet.ConditionalFormatting | ForEach-Object { [string]$_.Type }) | Should -Contain 'DataBar'
            @($worksheet.ConditionalFormatting | ForEach-Object { [string]$_.Type }) | Should -Contain 'ThreeColorScale'
            @($worksheet.ConditionalFormatting | ForEach-Object { [string]$_.Type }) | Should -Contain 'Top'
            @($worksheet.ConditionalFormatting | ForEach-Object { [string]$_.Type }) | Should -Contain 'Bottom'
        }
        finally {
            Close-ExcelPackage -ExcelPackage $excel -NoSave
        }
    }

    It "Accepts ImportExcel parameter aliases used by generated scripts" {
        $goodScript = @'
param()
$ws = $null
Set-ExcelRange -Address $ws.Cells["A1:B2"] -Bold
Set-ExcelRange -Address $ws.Cells["A1:B2"] -BorderAround Thin
Add-PivotTable -ExcelPackage $null -PivotTableName "Pivot" -SourceWorksheet $ws -SourceRange "A1:B2" -PivotRows "Region" -PivotData @{ Revenue = "Sum" }
'@

        $issues = @(Test-ExcelAiPowerShellCommandUsage -Script $goodScript)

        $issues.Count | Should -Be 0
    }

    It "Rejects generated scripts that use unsupported ImportExcel commands or parameters" {
        $badScript = @'
param([string]$SourcePath, [string]$OutputPath)
$excelPackage = New-ExcelPackage -Path $OutputPath
Set-ExcelRange -ExcelPackage $excelPackage -WorksheetName "Summary" -Range "A1"
Set-ExcelRange -Address $null -BorderAround
$ws.Cells["A1"].Style.Font.Color.SetColor("Blue")
Add-PivotTable -ExcelPackage $excelPackage -DestinationWorksheet "Pivot"
Save-ExcelPackage -ExcelPackage $excelPackage
'@

        $issues = @(Test-ExcelAiPowerShellCommandUsage -Script $badScript)

        $issues.Message | Should -Contain "Command 'New-ExcelPackage' was not found."
        $issues.Message | Should -Contain "Command 'Set-ExcelRange' does not have a '-ExcelPackage' parameter."
        $issues.Message | Should -Contain "Parameter '-BorderAround' for command 'Set-ExcelRange' requires an argument."
        $issues.Message | Should -Contain "Direct EPPlus SetColor calls are not supported in generated scripts. Use ImportExcel color parameters such as -BackgroundColor, -FontColor, or -BorderColor."
        $issues.Message | Should -Contain "Command 'Add-PivotTable' does not have a '-DestinationWorksheet' parameter."
        $issues.Message | Should -Contain "Command 'Save-ExcelPackage' was not found."
    }
}
