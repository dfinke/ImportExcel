Function Export-Excel {
    <# 
    .SYNOPSIS 
        Write objects and strings to an Excel worksheet.
 
    .DESCRIPTION 
        The Export-Excel cmdlet creates an Excel worksheet of the objects or strings you commit. This is done without using Microsoft Excel in the background but by using the .NET EPPLus.dll. You can also automate the creation of Pivot Tables and Charts.
 
    .PARAMETER Path 
        Specifies the path to the Excel file. This parameter is required.
 
    .PARAMETER TargetData 
        
    .PARAMETER WorksheetName
        Specifies the name of the worksheet in the Excel workbook.

    .PARAMETER Title
        Specifies the title used in the worksheet. The title is placed on the first line of the worksheet.

    .PARAMETER TitleFillPattern

    .PARAMETER TitleBold
        Sets the title to bold. By default the title is not bold.

    .PARAMETER TitleSize
        Specifies the size of the title. The default value is 22.

    .PARAMETER TitleBackgroundColor

    .PARAMETER PivotRows
        Specifies the rows in the pivot table.

    .PARAMETER PivotColumns
        Specifies the columns in the pivot table.

    .PARAMETER PivotData
        Specifies the source data in the pivot table.

    .PARAMETER Password
        Specifies the password to use to protect the Excel workbook from unauthorized access.

    .PARAMETER ChartType
        Specifies the type of chart to use. The default is a pie chart.

    .PARAMETER IncludePivotTable
        Adds a pivot table worksheet to the workbook. In data processing, a pivot table is a data summarization tool found in data visualization programs such as Excel worksheets. Among other functions, a pivot table can automatically sort, count total or give the average of the data stored in one table or spreadsheet, displaying the results in a second table showing the summarized data. Pivot tables are also useful for quickly creating unweighted cross tabulations. The user sets up and changes the summary's structure by dragging and dropping fields graphically in Excel or by using the paramaters 'PivotRows', 'PivotColumns', and 'PivotData'.

    .PARAMETER IncludePivotChart
        Has to be used together with 'IncludePivotTable' and adds an extra chart next to the pivot table.

    .PARAMETER AutoSize
        Adjusts the column width to the content of the cells. By default columns have a fixed width.

    .PARAMETER Show
        Opens the Excel file after creation, so you can view its content.

    .PARAMETER NoClobber
        Do not overwrite (replace the contents) of an existing worksheet. By default, if a file exists in the specified path, Export-Excel overwrites the worksheet without warning.

    .PARAMETER FreezeTopRow
        Freezes the first row of the Excel worksheet. This is convenient when working with lots of rows, so the the headers will always be visible when scrolling downwards in the worksheet.

    .PARAMETER AutoFilter
        Sets the auto filter on the first row. This allows you to view specific rows in an Excel spreadsheet, while hiding other rows in the worksheet. When the auto filter is added to the header row of a worksheet, a drop-down menu appears on each cell of the header row. This provides you with a number of filter options that can be used to specify which rows of the worksheet are to be displayed.

    .PARAMETER BoldTopRow
        Sets the top row of the worksheet to bold. By default the top row is not bold.

    .PARAMETER NoHeader
        Omits the header fields so the worksheet will not contain column headers.

    .PARAMETER RangeName

    .PARAMETER TableName
        Sets the content of the worksheet as a data table. Which makes it easier to sort, filter and maniuplate data in Excel.

    .PARAMETER ConditionalFormat

    .PARAMETER HideSheet
        Specifies which worksheets will be hidden in the workbook. By default, all worksheets are visible.
 
    .EXAMPLE
        Get-Service | Export-Excel .\Test.xlsx -WorksheetName 'Services' -TableName 'Services'
        Generates an Excel worksheet containing all the services on the system. The worksheet content will be presented in the Excel data table format for easy filtering, sorting and manipulation.

    .EXAMPLE
        Get-Service | Select-Object Status, Name, DisplayName | Export-Excel .\Test.xlsx -AutoSize -BoldTopRow -Show
        Generates an Excel worksheet containing all the services on the system. The worksheet will contain the headers 'Status', 'DisplayName' and 'Name' in bold. The column width will be adjusted to the cells content and the worksheet will be opened automatically once it's created.
        
        It will look like this:
        
        Sheet1:
        -------
        Status   Name               DisplayName
        Running  BITS               Background Intelligent Transfer Ser...
        Stopped  Browser            Computer Browser

    .EXAMPLE
        Get-Service | Export-Excel .\Test.xlsx -Show -NoHeader
        Generates an Excel worksheet containing all the services on the system. The worksheet will not contain any headers like 'Status', 'DisplayName' or 'Name' because we used the switch 'NoHeader'.
        
        It will look like this:
        
        Sheet1:
        -------
        Running  BITS               Background Intelligent Transfer Ser...
        Stopped  Browser            Computer Browser

    .EXAMPLE
        Get-Process | Export-Excel .\Test.xlsx -WorksheetName 'Processes'
        Get-Service | Export-Excel .\Test.xlsx -WorksheetName 'Services' -HideSheet 'Services'

        Creates an Excel workbook where only the worksheet 'Processes' is visible. The worksheet 'Services' is hidden.    

    .EXAMPLE
        Get-Process | Export-Excel .\Test.xlsx -WorksheetName Processes -IncludePivotTable -Show -PivotRows Company -PivotData PM
        Creates an Excel workbook containing two worksheets, one with a pivot table and one with the source data.

    .EXAMPLE
        $Params = @{
            Path              = '.\Test.xlsx'
            IncludePivotTable = $true
            PivotRows         = 'Status'
            PivotData         = @{Status='Count'}
            WorksheetName     = 'Services' 
            HideSheet         = 'Services'
            Show              = $true
        }
        Get-Service | Export-Excel @Params

        Creates two Excel worksheets, one with a pivot table named 'ServicesPivotTable' and one with the source worksheet named 'Services'. The last one will be hidden and the Excel file will be opened when the command finishes. You will only see the worksheet 'ServicesPivotTable' with the pivot table as the other one is hidden with the 'HideSheet' switch.

        It will look like this:

        ServicesPivotTable:
        -------------------
        Count of Status	
        Row Labels	| Total
        ----------- | -----
        Running	    | 87
        Stopped	    | 96
        Grand Total	| 183
    
    .EXAMPLE
        Get-Process | Export-Excel .\Test.xlsx -WorksheetName Processes -ChartType PieExploded3D -IncludePivotChart -IncludePivotTable -Show -PivotRows Company -PivotData PM
        Creates an Excel workbook containing two worksheets, one worksheet with a PieExploded3D chart and a pivot table, and one worksheet with the source data.

    .NOTES
        CHANGELOG
        2015/10/20 Added help text
        2015/10/20 Changed 'TitleBold' from [BOOL] to [Switch]
                   (Makes more sense then providing $true or $false)

    .LINK
        https://github.com/dfinke/ImportExcel

    #> 
    
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$Path,
        [Parameter(ValueFromPipeline=$true)]
        $TargetData,
        [String]$WorksheetName = 'Sheet1',
        [String]$Title,
        [OfficeOpenXml.Style.ExcelFillStyle]$TitleFillPattern = 'None',
        [Switch]$TitleBold,
        [Int]$TitleSize = 22,
        [System.Drawing.Color]$TitleBackgroundColor,
        [String[]]$PivotRows,
        [String[]]$PivotColumns,
        $PivotData,
        [String]$Password,
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = 'Pie',
        [Switch]$IncludePivotTable,
        [Switch]$IncludePivotChart,
        [Switch]$AutoSize,
        [Switch]$Show,
        [Switch]$NoClobber,
        [Switch]$FreezeTopRow,
        [Switch]$AutoFilter,
        [Switch]$BoldTopRow,
        [Switch]$NoHeader,
        [String]$RangeName,
        [String]$TableName,
        [Object[]]$ConditionalFormat,
        [String[]]$HideSheet
    )

    Begin {
        try {
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            if (Test-Path $path) {
                Write-Debug "File `"$Path`" already exists"
            }
            $pkg = New-Object OfficeOpenXml.ExcelPackage $Path

            $ws  = $pkg | Add-WorkSheet -WorksheetName $WorksheetName -NoClobber:$NoClobber

            foreach($format in $ConditionalFormat ) {
                $target = "Add$($format.Formatter)"
                $rule = ($ws.ConditionalFormatting).$target($format.Address, $format.IconType)
                $rule.Reverse = $format.Reverse
            }

            # Force at least one cell value
            $ws.Cells[1, 1].Value = ""

            $Row = 1
            if($Title) {
                $ws.Cells[$Row, 1].Value = $Title

                $ws.Cells[$Row, 1].Style.Font.Size = $TitleSize
                $ws.Cells[$Row, 1].Style.Font.Bold = $TitleBold
                $ws.Cells[$Row, 1].Style.Fill.PatternType = $TitleFillPattern
                if($TitleBackgroundColor) {
                    $ws.Cells[$Row, 1].Style.Fill.BackgroundColor.SetColor($TitleBackgroundColor)
                }

                $Row = 2
            }

        } Catch {
            if($AlreadyExists) {
                throw "$WorksheetName already exists."
            } else {
                throw $Error[0].Exception.Message
            }
        }

        $firstTimeThru = $true
        $isDataTypeValueType=$false
        $pattern = "string|bool|byte|char|decimal|double|float|int|long|sbyte|short|uint|ulong|ushort"
    }

    Process {
        if($firstTimeThru) {
            $firstTimeThru=$false
            $isDataTypeValueType = $TargetData.GetType().name -match "string|bool|byte|char|decimal|double|float|int|long|sbyte|short|uint|ulong|ushort"
        }

        if($isDataTypeValueType) {
            $ColumnIndex = 1

            $targetCell = $ws.Cells[$Row, $ColumnIndex]

            $r=$null
            $cellValue=$TargetData
            if([double]::tryparse($cellValue, [ref]$r)) {
                $targetCell.Value = $r
            } else {
                $targetCell.Value = $cellValue
            }

            switch ($TargetData.$Name) {
                {$_ -is [datetime]} {$targetCell.Style.Numberformat.Format = "m/d/yy h:mm"}
            }

            $ColumnIndex += 1
            $Row += 1

        } else {
            if(!$Header) {

                $ColumnIndex = 1
                $Header = $TargetData.psobject.properties.name

                if($NoHeader) {
                    # Don't push the headers to the spread sheet
                    $Row -= 1
                } else {
                    foreach ($Name in $Header) {
                        $ws.Cells[$Row, $ColumnIndex].Value = $name
                        $ColumnIndex += 1
                    }
                }
            }

            $Row += 1
            $ColumnIndex = 1

            foreach ($Name in $Header) {

                $targetCell = $ws.Cells[$Row, $ColumnIndex]

                $cellValue=$TargetData.$Name

                $r=$null
                if([double]::tryparse($cellValue, [ref]$r)) {
                    $targetCell.Value = $r
                } else {
                    $targetCell.Value = $cellValue
                }

                switch ($TargetData.$Name) {
                    {$_ -is [datetime]} {$targetCell.Style.Numberformat.Format = "m/d/yy h:mm"}
                }

                $ColumnIndex += 1
            }
        }
    }

    End {
        $startAddress=$ws.Dimension.Start.Address
        $dataRange="{0}:{1}" -f $startAddress, $ws.Dimension.End.Address
        Write-Debug "Data Range $dataRange"

        if (-not [string]::IsNullOrEmpty($RangeName)) {
            $ws.Names.Add($RangeName, $ws.Cells[$dataRange]) | Out-Null
        }
        if (-not [string]::IsNullOrEmpty($TableName)) {
            $ws.Tables.Add($ws.Cells[$dataRange], $TableName) | Out-Null
        }

        if($IncludePivotTable) {
            $pivotTableName = $WorksheetName + "PivotTable"
            $wsPivot = $pkg | Add-WorkSheet -WorksheetName $pivotTableName -NoClobber:$NoClobber

            $wsPivot.View.TabSelected = $true

            $pivotTableDataName=$WorksheetName + "PivotTableData"

            if($Title) {$startAddress="A2"}
            $pivotTable = $wsPivot.PivotTables.Add($wsPivot.Cells["A1"], $ws.Cells[$dataRange], $pivotTableDataName)

            if($PivotRows) {
                foreach ($Row in $PivotRows) {
                    $null=$pivotTable.RowFields.Add($pivotTable.Fields[$Row])
                }
            }

            if($PivotColumns) {
                foreach ($Column in $PivotColumns) {
                    $null=$pivotTable.ColumnFields.Add($pivotTable.Fields[$Column])
                }
            }

            if($PivotData) {
                if($PivotData -is [hashtable]) {
                    $PivotData.Keys | % {
                        $df=$pivotTable.DataFields.Add($pivotTable.Fields[$_])
                        $df.Function = $PivotData.$_
                    }
                } else {
                    foreach ($Item in $PivotData) {
                        $df=$pivotTable.DataFields.Add($pivotTable.Fields[$Item])
                        $df.Function = 'Count'
                    }
                }
            }

            if($IncludePivotChart) {
                $chart = $wsPivot.Drawings.AddChart("PivotChart", $ChartType, $pivotTable)
                $chart.SetPosition(1, 0, 6, 0)
                $chart.SetSize(600, 400)
            }
        }

        if($Password) { $ws.Protection.SetPassword($Password) }

        if($AutoFilter) {
            $ws.Cells[$dataRange].AutoFilter=$true
        }

        if($FreezeTopRow) {
            $ws.View.FreezePanes(2,1)
        }

        if($BoldTopRow) {
            $range=$ws.Dimension.Address -replace $ws.Dimension.Rows, "1"
            $ws.Cells[$range].Style.Font.Bold=$true
        }

        if($AutoSize) { $ws.Cells.AutoFitColumns() }

        #$pkg.Workbook.View.ActiveTab = $ws.SheetID

        foreach($Sheet in $HideSheet) {
            $pkg.Workbook.WorkSheets[$Sheet].Hidden="Hidden"
        }

        $pkg.Save()
        $pkg.Dispose()

        if($Show) {Invoke-Item $Path}
    }
}
