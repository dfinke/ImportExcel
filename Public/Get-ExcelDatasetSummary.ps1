function Get-ExcelDatasetSummary {
    <#
    .SYNOPSIS
        Profiles CSV/TSV and Excel files for AI-assisted report generation.

    .DESCRIPTION
        Reads one or more tabular sources from a CSV/TSV file or workbook and returns
        a compact summary suitable for prompts, planning, and deterministic report
        generation. The summary includes source names, row/column counts, inferred
        column types, roles, sample rows, and suggested number formats.

    .PARAMETER Path
        Path to a .csv, .tsv, .xlsx, .xlsm, .xltx, or .xltm file.

    .PARAMETER WorksheetName
        Optional worksheet names to profile when Path is an Excel workbook.

    .PARAMETER SampleSize
        Number of sample rows to include in the summary.

    .PARAMETER AnalyzeRows
        Maximum number of rows per source to inspect for type inference.

    .PARAMETER Delimiter
        Delimiter for .csv files. .tsv files always use a tab delimiter.

    .PARAMETER AsJson
        Returns the summary as JSON instead of PowerShell objects.

    .EXAMPLE
        Get-ExcelDatasetSummary -Path .\sales.csv

    .EXAMPLE
        Get-ExcelDatasetSummary -Path .\workbook.xlsx -WorksheetName Sales,Inventory -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [string]$Path,

        [string[]]$WorksheetName,

        [ValidateRange(0, 100)]
        [int]$SampleSize = 5,

        [ValidateRange(1, [int]::MaxValue)]
        [int]$AnalyzeRows = 1000,

        [char]$Delimiter = ',',

        [switch]$AsJson,

        [switch]$Compress
    )

    process {
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        $sources = @(Get-ExcelAiSourceData -Path $resolvedPath -WorksheetName $WorksheetName -Delimiter $Delimiter)
        $sourceSummaries = @()

        foreach ($source in $sources) {
            $rows = @($source.Data)
            $rowsToAnalyze = @($rows | Select-Object -First $AnalyzeRows)
            $columns = @(Get-ExcelAiColumnSummary -Rows $rowsToAnalyze)

            $sourceSummaries += [pscustomobject][ordered]@{
                Name = $source.Name
                Kind = $source.Kind
                RowCount = $rows.Count
                ColumnCount = $columns.Count
                ColumnNames = @($columns | Select-Object -ExpandProperty Name)
                ColumnSummaries = $columns
                SampleRows = @($rows | Select-Object -First $SampleSize)
            }
        }

        $summary = [pscustomobject][ordered]@{
            Path = $resolvedPath
            FileName = [System.IO.Path]::GetFileName($resolvedPath)
            FileType = [System.IO.Path]::GetExtension($resolvedPath).TrimStart('.')
            GeneratedAt = Get-Date
            Sources = $sourceSummaries
        }

        if ($AsJson) {
            return $summary | ConvertTo-Json -Depth 20 -Compress:$Compress
        }

        return $summary
    }
}

