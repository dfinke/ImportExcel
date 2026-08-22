# ImportExcel MCP Server Plan

## Feasibility & Approach

**Bottom line: High feasibility, low effort (~2-3 hours).** The hard part (MCP protocol, type coercion, JSON serialization) is already done in `PSMCPServer.ps1`. You just need a thin wrapper script.

---

## The Challenge with Direct Registration

`PSMCPServer.ps1` auto-registers functions by inspecting their parameter sets. `Import-Excel` and `Export-Excel` have problems for direct registration:

1. **Multiple parameter sets** — `Register-MCPTool` only handles one set (index 0)
2. **Export-Excel's `$InputObject`** is pipeline-based — MCP calls pass arguments, not pipelines
3. **Complex parameter types** — `ExcelChartDefinition`, `PivotTableDefinition`, etc. won't translate to JSON schema meaningfully
4. **Too many params** — Export-Excel has ~60+ params; an LLM shouldn't be handed that wall

---

## Recommended Approach: Thin Wrapper Functions

Create an `ImportExcel-MCP.ps1` wrapper script that:
1. Imports the `ImportExcel` module
2. Defines 4-5 **simplified wrapper functions** with clean comment-based help and flat parameter sets
3. Calls `Start-PSMCPServer` with those wrappers

---

## Suggested MCP Tools (initial set)

| Tool | Wraps | Key params for LLM |
|------|-------|-------------------|
| `Read-ExcelFile` | `Import-Excel` | `Path`, `WorksheetName`, `StartRow`, `EndRow` |
| `Export-DataToExcel` | `Export-Excel` | `Path`, `JsonData` (string), `WorksheetName`, `AutoSize`, `AutoFilter`, `BoldTopRow` |
| `Get-ExcelSheets` | `Get-ExcelSheetInfo` | `Path` |
| `Get-ExcelSchema` | `Get-ExcelFileSchema` | `Path` |

`Export-DataToExcel` would accept a **JSON string** of data, `ConvertFrom-Json` it, then pipe to `Export-Excel`. That bridges the pipeline gap cleanly.

---

## The Killer Use Case

An LLM could:
1. Call `Get-ExcelSchema` → discover column names
2. Call `Read-ExcelFile` → get the data as JSON
3. Transform/analyze it
4. Call `Export-DataToExcel` → write results back to a new sheet

That's a complete Excel data pipeline driven by natural language.

---

## Effort Estimate

- **~30 min**: Write the 4 wrapper functions with good `.SYNOPSIS`/`.PARAMETER` help
- **~30 min**: Test each tool via the MCP server manually
- **~30 min**: Wire up `claude_desktop_config.json` entry and verify end-to-end

## Reference

- MCP server framework: `D:\mygit\PowerShellAIAssistant-ScratchPad\PowerShell-Codex-CLI\PSMCPServer.ps1`
- ImportExcel public functions: `D:\mygit\ImportExcel\Public\`
