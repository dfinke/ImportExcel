---
name: importexcel-reporting
description: Use ImportExcel to inspect, summarize, and generate Excel reports from data sources.
---

When asked to work with Excel or CSV data:

1. Prefer the ImportExcel PowerShell module for workbook inspection, export, and report generation.
2. Use the existing AI-aware flow when the user wants a natural-language report:
   - Invoke-ExcelPrompt
   - Invoke-ExcelAgent
   - New-ExcelReportScript
3. Use the deterministic fallback path when the model is unavailable or the user wants a reliable offline workflow.
4. Keep the output practical: workbook path, summary of sheets and columns, and reusable PowerShell script if requested.
5. If the user wants a simple, safe path, suggest:
   - Get-ExcelDatasetSummary
   - New-ExcelReportPlan
   - Invoke-ExcelReportPlan
