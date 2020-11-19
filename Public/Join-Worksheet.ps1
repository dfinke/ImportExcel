function Join-Worksheet {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param (
        [Parameter(ParameterSetName = "Default", Position = 0)]
        [Parameter(ParameterSetName = "Table"  , Position = 0)]
        [String]$Path  ,
        [Parameter(Mandatory = $true, ParameterSetName = "PackageDefault")]
        [Parameter(Mandatory = $true, ParameterSetName = "PackageTable")]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        $WorksheetName = 'Combined',
        [switch]$Clearsheet,
        [switch]$NoHeader,
        [string]$FromLabel = "From" ,
        [switch]$LabelBlocks,
        [Switch]$AutoSize,
        [Switch]$FreezeTopRow,
        [Switch]$FreezeFirstColumn,
        [Switch]$FreezeTopRowFirstColumn,
        [Int[]]$FreezePane,
        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'PackageDefault')]
        [Switch]$AutoFilter,
        [Switch]$BoldTopRow,
        [switch]$HideSource,
        [String]$Title,
        [OfficeOpenXml.Style.ExcelFillStyle]$TitleFillPattern = 'Solid',
        $TitleBackgroundColor,
        [Switch]$TitleBold,
        [Int]$TitleSize = 22,
        [Hashtable]$PivotTableDefinition,
        [Object[]]$ExcelChartDefinition,
        [Object[]]$ConditionalFormat,
        [Object[]]$ConditionalText,
        [switch]$AutoNameRange,
        [ValidateScript( {
                if (-not $_) {  throw 'RangeName is null or empty.'  }
                elseif ($_[0] -notmatch '[a-z]') { throw 'RangeName starts with an invalid character.'  }
                else { $true }
        })]
        [String]$RangeName,
        [ValidateScript( {
                if (-not $_) {  throw 'Tablename is null or empty.'  }
                elseif ($_[0] -notmatch '[a-z]') { throw 'Tablename starts with an invalid character.'  }
                else { $true }
        })]
        [Parameter(ParameterSetName = 'Table'        , Mandatory = $true)]
        [Parameter(ParameterSetName = 'PackageTable' , Mandatory = $true)]
        [String]$TableName,
        [Parameter(ParameterSetName = 'Table')]
        [Parameter(ParameterSetName = 'PackageTable')]
        [OfficeOpenXml.Table.TableStyles]$TableStyle = 'Medium6',
        [switch]$ReturnRange,
        [switch]$Show,
        [switch]$PassThru
    )
    #region get target worksheet, select it and move it to the end.
    if ($Path -and -not $ExcelPackage) {$ExcelPackage = Open-ExcelPackage -path $Path  }
    $destinationSheet = Add-Worksheet -ExcelPackage $ExcelPackage -WorksheetName $WorksheetName -ClearSheet:$Clearsheet
    foreach ($w in $ExcelPackage.Workbook.Worksheets) {$w.view.TabSelected = $false}
    $destinationSheet.View.TabSelected = $true
    $ExcelPackage.Workbook.Worksheets.MoveToEnd($WorksheetName)
    #row to insert at will be 1 on a blank sheet and lastrow + 1 on populated one
    $row = (1 + $destinationSheet.Dimension.End.Row )
    #endregion

    #region Setup title and header rows
    #Title parameters work as they do in Export-Excel .
    if ($row -eq 1 -and $Title) {
        $destinationSheet.Cells[1, 1].Value = $Title
        $destinationSheet.Cells[1, 1].Style.Font.Size = $TitleSize
        if ($TitleBold) {$destinationSheet.Cells[1, 1].Style.Font.Bold = $True }
        #Can only set TitleBackgroundColor if TitleFillPattern is something other than None.
        if ($TitleBackgroundColor -AND ($TitleFillPattern -ne 'None')) {
            if ($TitleBackgroundColor -is [string])         {$TitleBackgroundColor = [System.Drawing.Color]::$TitleBackgroundColor }
            $destinationSheet.Cells[1, 1].Style.Fill.PatternType = $TitleFillPattern
            $destinationSheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor($TitleBackgroundColor)
        }
        elseif ($TitleBackgroundColor) { Write-Warning "Title Background Color ignored. You must set the TitleFillPattern parameter to a value other than 'None'. Try 'Solid'." }
        $row = 2
    }

    if (-not $noHeader) {
        #Assume every row has titles in row 1, copy row 1 from first sheet to new sheet.
        $destinationSheet.Select("A$row")
        $ExcelPackage.Workbook.Worksheets[1].cells["1:1"].Copy($destinationSheet.SelectedRange)
        #fromlabel can't be an empty string
        if ($FromLabel ) {
            #Add a column which says where the data comes from.
            $fromColumn = ($destinationSheet.Dimension.Columns + 1)
            $destinationSheet.Cells[$row, $fromColumn].Value = $FromLabel
        }
        $row += 1
    }
    #endregion

    foreach ($i in 1..($ExcelPackage.Workbook.Worksheets.Count - 1) ) {
        $sourceWorksheet = $ExcelPackage.Workbook.Worksheets[$i]
        #Assume row one is titles, so data itself starts at A2.
        if ($NoHeader) {$sourceRange = $sourceWorksheet.Dimension.Address}
        else {$sourceRange = $sourceWorksheet.Dimension.Address -replace "A1:", "A2:"}
        #Position insertion point/
        $destinationSheet.Select("A$row")
        if ($LabelBlocks) {
            $destinationSheet.Cells[$row, 1].value = $sourceWorksheet.Name
            $destinationSheet.Cells[$row, 1].Style.Font.Bold = $true
            $destinationSheet.Cells[$row, 1].Style.Font.Size += 2
            $row += 1
        }
        $destinationSheet.Select("A$row")

        #And finally we're ready to copy the data.
        $sourceWorksheet.Cells[$sourceRange].Copy($destinationSheet.SelectedRange)
        #Fill in column saying where data came from.
        if ($fromColumn) { $row..$destinationSheet.Dimension.Rows | ForEach-Object {$destinationSheet.Cells[$_, $fromColumn].Value = $sourceWorksheet.Name} }
        #Update where next insertion will go.
        $row = $destinationSheet.Dimension.Rows + 1
        if ($HideSource) {$sourceWorksheet.Hidden = [OfficeOpenXml.eWorkSheetHidden]::Hidden}
    }

    #We accept a bunch of parameters work to pass on to Export-excel ( Autosize, Autofilter, boldtopRow Freeze ); if we have any of those call Export-excel otherwise close the package here.
    $params = @{} + $PSBoundParameters
    'Path', 'Clearsheet', 'NoHeader', 'FromLabel', 'LabelBlocks', 'HideSource',
    'Title', 'TitleFillPattern', 'TitleBackgroundColor', 'TitleBold', 'TitleSize' | ForEach-Object {$null = $params.Remove($_)}
    if ($params.Keys.Count) {
        if ($Title) { $params.StartRow = 2}
        $params.WorksheetName = $WorksheetName
        $params.ExcelPackage = $ExcelPackage
        Export-Excel @Params
    }
    else {
        Close-ExcelPackage -ExcelPackage $ExcelPackage
        $ExcelPackage.Dispose()
        $ExcelPackage = $null
    }
}