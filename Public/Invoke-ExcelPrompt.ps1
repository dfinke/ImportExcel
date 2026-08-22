function Invoke-ExcelPrompt {
    <#
    .SYNOPSIS
        Creates an Excel report from a natural-language prompt.

    .DESCRIPTION
        Profiles a CSV/TSV/workbook, builds a report plan with New-ExcelReportPlan,
        and executes that plan with ImportExcel. Use -NoAI to create a useful
        deterministic starter report without a model call.

    .PARAMETER Path
        Source CSV/TSV/workbook path.

    .PARAMETER Prompt
        Natural-language report request.

    .PARAMETER OutputPath
        Destination workbook. Defaults to "<source name>-report.xlsx" beside the source.

    .PARAMETER Model
        PSAISuite model identifier.

    .PARAMETER PSAISuitePath
        Optional path to PSAISuite.psd1.

    .PARAMETER NoAI
        Builds the deterministic starter plan instead of calling PSAISuite.

    .PARAMETER Force
        Overwrites an existing output workbook.

    .PARAMETER Show
        Opens the workbook after it is created.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Prompt,

        [string]$OutputPath,

        [string[]]$WorksheetName,

        [char]$Delimiter = ',',

        [string]$Model = $(if ($env:IMPORTEXCEL_AI_MODEL) { $env:IMPORTEXCEL_AI_MODEL } elseif ($env:PSAISUITE_DEFAULT_MODEL) { $env:PSAISUITE_DEFAULT_MODEL } else { 'openai:gpt-5.2' }),

        [string]$PSAISuitePath,

        [switch]$NoAI,

        [switch]$FallbackToDefault,

        [switch]$Force,

        [switch]$Show
    )

    process {
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        if (-not $OutputPath) {
            $directory = Split-Path -Path $resolvedPath -Parent
            $fileName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath) + '-report.xlsx'
            $OutputPath = Join-Path -Path $directory -ChildPath $fileName
        }

        $summary = Get-ExcelDatasetSummary -Path $resolvedPath -WorksheetName $WorksheetName -Delimiter $Delimiter
        $plan = New-ExcelReportPlan -DatasetSummary $summary -Prompt $Prompt -Model $Model -PSAISuitePath $PSAISuitePath -NoAI:$NoAI -FallbackToDefault:$FallbackToDefault
        Invoke-ExcelReportPlan -Plan $plan -Path $OutputPath -SourcePath $resolvedPath -DatasetSummary $summary -WorksheetName $WorksheetName -Delimiter $Delimiter -Force:$Force -Show:$Show
    }
}

