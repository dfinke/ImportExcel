<#
.SYNOPSIS
    ImportExcel MCP Server — exposes Import-Excel and Export-Excel as MCP tools.

.DESCRIPTION
    A stdio MCP server wrapping the ImportExcel PowerShell module.
    Exposes four tools: Read-ExcelFile, Export-DataToExcel, Get-ExcelSheets, Get-ExcelSchema.

.EXAMPLE
    # Claude Desktop config (claude_desktop_config.json):
    # "importexcel": {
    #   "command": "powershell",
    #   "args": ["-NoProfile", "-File", "C:\\path\\to\\ImportExcel-MCP.ps1"]
    # }
#>

# Dot-source the PSMCPServer framework
. "D:\mygit\PowerShellAIAssistant-ScratchPad\PowerShell-Codex-CLI\PSMCPServer.ps1"

# Import the ImportExcel module
Import-Module "$PSScriptRoot\ImportExcel.psd1" -ErrorAction Stop

# ─────────────────────────────────────────────────────────────
# Wrapper tools
# ─────────────────────────────────────────────────────────────

function Read-ExcelFile {
    <#
    .SYNOPSIS
        Reads data from an Excel file and returns it as JSON.

    .PARAMETER Path
        Full path to the Excel file (.xlsx).

    .PARAMETER WorksheetName
        Name of the worksheet to read. Defaults to the first sheet.

    .PARAMETER StartRow
        Row number where data begins (default 1).

    .PARAMETER EndRow
        Last row to read. Omit to read all rows.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [string]$WorksheetName,

        [int]$StartRow = 1,

        [int]$EndRow
    )

    $params = @{ Path = $Path; StartRow = $StartRow }
    if ($WorksheetName) { $params['WorksheetName'] = $WorksheetName }
    if ($EndRow)        { $params['EndRow']        = $EndRow        }

    Import-Excel @params
}

function Export-DataToExcel {
    <#
    .SYNOPSIS
        Writes JSON data to an Excel file.

    .PARAMETER Path
        Full path for the output Excel file (.xlsx).

    .PARAMETER JsonData
        JSON string representing an array of objects to write.

    .PARAMETER WorksheetName
        Name of the worksheet. Defaults to Sheet1.

    .PARAMETER AutoSize
        Auto-size columns to fit content.

    .PARAMETER AutoFilter
        Add an AutoFilter to the header row.

    .PARAMETER BoldTopRow
        Bold the header row.

    .PARAMETER Append
        Append data to an existing sheet instead of overwriting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$JsonData,

        [string]$WorksheetName = 'Sheet1',

        [switch]$AutoSize,

        [switch]$AutoFilter,

        [switch]$BoldTopRow,

        [switch]$Append
    )

    $data = $JsonData | ConvertFrom-Json

    $params = @{
        Path          = $Path
        WorksheetName = $WorksheetName
        AutoSize      = $AutoSize
        AutoFilter    = $AutoFilter
        BoldTopRow    = $BoldTopRow
        Append        = $Append
    }

    $data | Export-Excel @params

    [PSCustomObject]@{ Status = 'OK'; Path = $Path; Rows = @($data).Count }
}

function Get-ExcelSheets {
    <#
    .SYNOPSIS
        Lists all worksheets in an Excel file.

    .PARAMETER Path
        Full path to the Excel file (.xlsx).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    Get-ExcelSheetInfo -Path $Path
}

function Get-ExcelSchema {
    <#
    .SYNOPSIS
        Returns the column names (schema) for each worksheet in an Excel file.

    .PARAMETER Path
        Full path to the Excel file (.xlsx).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    Get-ExcelFileSchema -Path $Path
}

# ─────────────────────────────────────────────────────────────
# Start the server
# ─────────────────────────────────────────────────────────────

Start-PSMCPServer `
    -Functions 'Read-ExcelFile', 'Export-DataToExcel', 'Get-ExcelSheets', 'Get-ExcelSchema' `
    -ServerName 'importexcel-mcp' `
    -ServerVersion '1.0.0'
