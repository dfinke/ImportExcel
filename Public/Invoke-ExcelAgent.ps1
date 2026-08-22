function Invoke-ExcelAgent {
    <#
    .SYNOPSIS
        Generates and optionally runs reusable ImportExcel automation from a prompt.

    .DESCRIPTION
        Creates a PowerShell report script with New-ExcelReportScript and runs it to
        produce the workbook unless -ScriptOnly is supplied. This is the AI-first
        entry point: the workbook is useful, and the automation script remains as a
        reusable artifact.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Prompt,

        [string]$OutputPath,

        [string]$ScriptPath,

        [string[]]$WorksheetName,

        [char]$Delimiter = ',',

        [string]$Model = $(if ($env:IMPORTEXCEL_AI_MODEL) { $env:IMPORTEXCEL_AI_MODEL } elseif ($env:PSAISUITE_DEFAULT_MODEL) { $env:PSAISUITE_DEFAULT_MODEL } else { 'openai:gpt-4o-mini' }),

        [string]$PSAISuitePath,

        [switch]$NoAI,

        [switch]$FallbackToDefault,

        [switch]$ScriptOnly,

        [switch]$Force,

        [switch]$Show
    )

    process {
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        if (-not $OutputPath) {
            $directory = Split-Path -Path $resolvedPath -Parent
            $fileName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath) + '-agent-report.xlsx'
            $OutputPath = Join-Path -Path $directory -ChildPath $fileName
        }
        $resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)

        if (-not $ScriptPath) {
            $scriptFileName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedOutputPath) + '.generated.ps1'
            $ScriptPath = Join-Path -Path (Split-Path -Path $resolvedOutputPath -Parent) -ChildPath $scriptFileName
        }

        $scriptResult = New-ExcelReportScript `
            -Path $resolvedPath `
            -Prompt $Prompt `
            -OutputPath $resolvedOutputPath `
            -ScriptPath $ScriptPath `
            -WorksheetName $WorksheetName `
            -Delimiter $Delimiter `
            -Model $Model `
            -PSAISuitePath $PSAISuitePath `
            -NoAI:$NoAI `
            -FallbackToDefault:$FallbackToDefault `
            -Force:$Force

        $runResult = $null
        if (-not $ScriptOnly) {
            if ((Test-Path -LiteralPath $resolvedOutputPath) -and -not $Force) {
                throw "Output workbook already exists: $resolvedOutputPath. Use -Force to overwrite it."
            }

            try {
                $runResult = & $scriptResult.ScriptPath -SourcePath $resolvedPath -OutputPath $resolvedOutputPath -Show:$Show -Force:$Force
            }
            catch {
                if ($FallbackToDefault -and $scriptResult.UsedAI) {
                    Write-Warning "Falling back to the deterministic report script because the generated script failed while running: $_"
                    $scriptResult = New-ExcelReportScript `
                        -Path $resolvedPath `
                        -Prompt $Prompt `
                        -OutputPath $resolvedOutputPath `
                        -ScriptPath $ScriptPath `
                        -WorksheetName $WorksheetName `
                        -Delimiter $Delimiter `
                        -NoAI `
                        -Force
                    $runResult = & $scriptResult.ScriptPath -SourcePath $resolvedPath -OutputPath $resolvedOutputPath -Show:$Show -Force:$Force
                }
                else {
                    throw
                }
            }
        }

        [pscustomobject][ordered]@{
            Path = $resolvedOutputPath
            ScriptPath = $scriptResult.ScriptPath
            SourcePath = $resolvedPath
            UsedAI = $scriptResult.UsedAI
            Model = $scriptResult.Model
            ScriptOnly = [bool]$ScriptOnly
            Result = $runResult
        }
    }
}
