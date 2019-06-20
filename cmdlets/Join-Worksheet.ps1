function Join-Worksheet {
    <#
      .SYNOPSIS
        Combines data on all the sheets in an Excel worksheet onto a single sheet.
      .DESCRIPTION
        Join-Worksheet can work in two main ways, either
        Combining data which has the same layout from many pages into one, or
        combining pages which have nothing in common.
        In the former case the header row is copied from the first sheet and,
        by default, each row of data is labelled with the name of the sheet it came from.
        In the latter case -NoHeader is specified, and each copied block can have the
        sheet it came from placed above it as a title.
      .EXAMPLE
      >
      PS> foreach ($computerName in @('Server1', 'Server2', 'Server3', 'Server4')) {
      Get-Service -ComputerName $computerName |  Select-Object -Property Status, Name, DisplayName, StartType |
                Export-Excel -Path .\test.xlsx -WorkSheetname $computerName -AutoSize
      }
      PS> $ptDef =New-PivotTableDefinition -PivotTableName "Pivot1" -SourceWorkSheet "Combined" -PivotRows "Status" -PivotFilter "MachineName" -PivotData @{Status='Count'} -IncludePivotChart -ChartType BarClustered3D
      PS> Join-Worksheet -Path .\test.xlsx -WorkSheetName combined -FromLabel "MachineName" -HideSource  -AutoSize -FreezeTopRow -BoldTopRow  -PivotTableDefinition $pt -Show

      The foreach command gets the services running on four servers and exports each
      to its own page in Test.xlsx.
      $PtDef=  creates a definition for a PivotTable.
      The Join-Worksheet command uses the same file and merges the results into a sheet
      named "Combined". It sets a column header of "Machinename", this column will
      contain the name of the sheet the data was copied from; after copying the data
      to the sheet "Combined", the other sheets will be hidden.
      Join-Worksheet finishes by calling Export-Excel to AutoSize cells, freeze the
      top row and make it bold and add thePivotTable.

      .EXAMPLE
      >
      PS> Get-WmiObject -Class win32_logicaldisk | Select-Object -Property DeviceId,VolumeName, Size,Freespace |
                Export-Excel -Path "$env:computerName.xlsx" -WorkSheetname Volumes -NumberFormat "0,000"
      PS> Get-NetAdapter  | Select-Object Name,InterfaceDescription,MacAddress,LinkSpeed |
                Export-Excel -Path "$env:COMPUTERNAME.xlsx" -WorkSheetname NetAdapter
      PS> Join-Worksheet -Path "$env:COMPUTERNAME.xlsx"  -WorkSheetName Summary -Title "Summary" -TitleBold -TitleSize 22 -NoHeader -LabelBlocks -AutoSize -HideSource -show

      The first two commands get logical-disk and network-card information; each type
      is exported to its own sheet in a workbook.
      The Join-Worksheet command copies both onto a page named "Summary". Because
      the data is dissimilar, -NoHeader is specified, ensuring the whole of each
      page is copied. Specifying -LabelBlocks causes each sheet's name to become
      a title on the summary page above the copied data. The source data is
      hidden, a title is added in 22 point boldface and the columns are sized
      to fit the data.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param (
        # Path to a new or existing .XLSX file.
        [Parameter(ParameterSetName = "Default", Position = 0)]
        [Parameter(ParameterSetName = "Table"  , Position = 0)]
        [String]$Path  ,
        # An object representing an Excel Package - either from Open-Excel Package or  specifying -PassThru to Export-Excel.
        [Parameter(Mandatory = $true, ParameterSetName = "PackageDefault")]
        [Parameter(Mandatory = $true, ParameterSetName = "PackageTable")]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        # The name of a sheet within the workbook where the other sheets will be joined together - "Combined" by default.
        $WorkSheetName = 'Combined',
        # If specified ,any pre-existing target for the joined data will be deleted and re-created; otherwise data will be appended on this sheet.
        [switch]$Clearsheet,
        #Join-Worksheet assumes each sheet has identical headers and the headers should be copied to the target sheet, unless -NoHeader is specified.
        [switch]$NoHeader,
        #If -NoHeader is NOT specified, then rows of data will be labeled with the name of the sheet they came from. FromLabel is the header for this column. If it is null or empty, the labels will be omitted.
        [string]$FromLabel = "From" ,
        #If specified, the copied blocks of data will have the name of the sheet they were copied from inserted above them as a title.
        [switch]$LabelBlocks,
        #Sets the width of the Excel columns to display all the data in their cells.
        [Switch]$AutoSize,
        #Freezes headers etc. in the top row.
        [Switch]$FreezeTopRow,
        #Freezes titles etc. in the left column.
        [Switch]$FreezeFirstColumn,
        #Freezes top row and left column (equivalent to Freeze pane 2,2 ).
        [Switch]$FreezeTopRowFirstColumn,
        # Freezes panes at specified coordinates (in the form  RowNumber , ColumnNumber).
        [Int[]]$FreezePane,
        #Enables the Excel filter on the headers of the combined sheet.
        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'PackageDefault')]
        [Switch]$AutoFilter,
        #Makes the top row boldface.
        [Switch]$BoldTopRow,
        #If specified, hides the sheets that the data is copied from.
        [switch]$HideSource,
        #Text of a title to be placed in Cell A1.
        [String]$Title,
        #Sets the fill pattern for the title cell.
        [OfficeOpenXml.Style.ExcelFillStyle]$TitleFillPattern = 'Solid',
        #Sets the cell background color for the title cell.
        $TitleBackgroundColor,
        #Sets the title in boldface type.
        [Switch]$TitleBold,
        #Sets the point size for the title.
        [Int]$TitleSize = 22,
        #Hashtable(s) with Sheet PivotRows, PivotColumns, PivotData, IncludePivotChart and ChartType values to specify a definition for one or morePivotTable(s).
        [Hashtable]$PivotTableDefinition,
        #A hashtable containing ChartType, Title, NoLegend, ShowCategory, ShowPercent, Yrange, Xrange and SeriesHeader for one or more [non-pivot] charts.
        [Object[]]$ExcelChartDefinition,
        #One or more conditional formatting rules defined with New-ConditionalFormattingIconSet.
        [Object[]]$ConditionalFormat,
        #Applies a Conditional formatting rule defined with New-ConditionalText. When specific conditions are met the format is applied
        [Object[]]$ConditionalText,
        #Makes each column a named range.
        [switch]$AutoNameRange,
        #Makes the data in the worksheet a named range.
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
        # Makes the data in the worksheet a table with a name and applies a style to it. Name must not contain spaces.
        [String]$TableName,
        [Parameter(ParameterSetName = 'Table')]
        [Parameter(ParameterSetName = 'PackageTable')]
        #Selects the style for the named table - defaults to "Medium6".
        [OfficeOpenXml.Table.TableStyles]$TableStyle = 'Medium6',
        #If specified, returns the range of cells in the combined sheet, in the format "A1:Z100".
        [switch]$ReturnRange,
        #Opens the Excel file immediately after creation. Convenient for viewing the results instantly without having to search for the file first.
        [switch]$Show,
        #If specified, an object representing the unsaved Excel package will be returned, it then needs to be saved.
        [switch]$PassThru
    )
    #region get target worksheet, select it and move it to the end.
    if ($Path -and -not $ExcelPackage) {$ExcelPackage = Open-ExcelPackage -path $Path  }
    $destinationSheet = Add-WorkSheet -ExcelPackage $ExcelPackage -WorkSheetname $WorkSheetName -ClearSheet:$Clearsheet
    foreach ($w in $ExcelPackage.Workbook.Worksheets) {$w.view.TabSelected = $false}
    $destinationSheet.View.TabSelected = $true
    $ExcelPackage.Workbook.Worksheets.MoveToEnd($WorkSheetName)
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

    #We accept a bunch of parameters work to pass on to Export-excel ( Autosize, Autofilter, boldtopRow Freeze ); if we have any of those call export-excel otherwise close the package here.
    $params = @{} + $PSBoundParameters
    'Path', 'Clearsheet', 'NoHeader', 'FromLabel', 'LabelBlocks', 'HideSource',
    'Title', 'TitleFillPattern', 'TitleBackgroundColor', 'TitleBold', 'TitleSize' | ForEach-Object {$null = $params.Remove($_)}
    if ($params.Keys.Count) {
        if ($Title) { $params.StartRow = 2}
        $params.WorkSheetName = $WorkSheetName
        $params.ExcelPackage = $ExcelPackage
        Export-Excel @Params
    }
    else {
        Close-ExcelPackage -ExcelPackage $ExcelPackage
        $ExcelPackage.Dispose()
        $ExcelPackage = $null
    }
}