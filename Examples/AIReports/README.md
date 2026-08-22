# AI-assisted report examples

These scripts demonstrate the agentic report-planning layer:

- `01-CreateStarterReport.ps1` creates a sales CSV and builds a report with `Invoke-ExcelPrompt -NoAI`.
- `02-InspectPlanAndRender.ps1` shows the lower-level flow: summarize data, create a plan, save the JSON plan, then render it.
- `03-UsePSAISuitePlanner.ps1` asks PSAISuite to create the plan, with `-FallbackToDefault` so it still renders a starter report when PSAISuite or API keys are unavailable.
- `04-ReportFromExistingWorkbook.ps1` creates a multi-sheet workbook, then reads it back as the source for an AI-assisted report.
- `05-GenerateReusableScript.ps1` generates a reusable PowerShell script and runs it to build the workbook. It hides fallback warnings by default for a cleaner demo; add `-ShowFallbackWarnings` when tuning the AI-generated script path.

Run any script from PowerShell:

```powershell
.\Examples\AIReports\01-CreateStarterReport.ps1 -Show
```

The examples write inputs and outputs under `$env:TEMP\ImportExcelAIExamples`.
