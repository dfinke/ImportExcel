function Export-Excel {
    <#
        .SYNOPSIS
            Export data to an Excel worksheet.
        .DESCRIPTION
            Export data to an Excel file and where possible try to convert numbers so Excel recognizes them as numbers instead of text. After all. Excel is a spreadsheet program used for number manipulation and calculations. In case the number conversion is not desired, use the parameter '-NoNumberConversion *'.
        .PARAMETER Path
            Path to a new or existing .XLSX file.
        .PARAMETER  ExcelPackage
            An object representing an Excel Package - usually this is returned by specifying -Passthru allowing multiple commands to work on the same Workbook without saving and reloading each time.
        .PARAMETER WorkSheetName
            The name of a sheet within the workbook - "Sheet1" by default.
        .PARAMETER ClearSheet
            If specified Export-Excel will remove any existing worksheet with the selected name. The Default behaviour is to overwrite cells in this sheet as needed (but leaving non-overwritten ones in place).
        .PARAMETER Append
            If specified data will be added to the end of an existing sheet, using the same column headings.
        .PARAMETER TargetData
            Data to insert onto the worksheet - this is often provided from the pipeline.
        .PARAMETER ExcludeProperty
            Specifies properties which may exist in the target data but should not be placed on the worksheet.
        .PARAMETER Title
            Text of a title to be placed in Cell A1.
        .PARAMETER TitleBold
            Sets the title in boldface type.
        .PARAMETER TitleSize
            Sets the point size for the title.
        .PARAMETER TitleBackgroundColor
            Sets the cell background color for the title cell.
        .PARAMETER TitleFillPattern
            Sets the fill pattern for the title cell.
        .PARAMETER Password
            Sets password protection on the workbook.
        .PARAMETER IncludePivotTable
            Adds a Pivot table using the data in the worksheet.
        .PARAMETER PivotRows
            Name(s) columns from the spreadhseet which will provide the row name(s) in the pivot table.
        .PARAMETER PivotColumns
            Name(s) columns from the spreadhseet which will provide the Column name(s) in the pivot table.
        .PARAMETER PivotData
            Hash table in the form ColumnName = Average|Count|CountNums|Max|Min|Product|None|StdDev|StdDevP|Sum|Var|VarP to provide the data in the Pivot table.
        .PARAMETER PivotTableDefinition,
            HashTable(s) with Sheet PivotTows, PivotColumns, PivotData, IncludePivotChart and ChartType values to make it easier to specify a definition or multiple Pivots.
        .PARAMETER IncludePivotChart,
             Include a chart with the Pivot table - implies Include Pivot Table.
        .PARAMETER NoLegend
            Exclude the legend from the pivot chart.
        .PARAMETER ShowCategory
            Add category labels to the pivot chart.
        .PARAMETER ShowPercent
            Add Percentage labels to the pivot chart.
        .PARAMETER ConditionalText
            Applies a 'Conditional formatting rule' in Excel on all the cells. When specific conditions are met a rule is triggered.
        .PARAMETER NoNumberConversion
            By default we convert all values to numbers if possible, but this isn't always desirable. NoNumberConversion allows you to add exceptions for the conversion. Wildcards (like '*') are allowed.
        .PARAMETER BoldTopRow
            Makes the top Row boldface.
        .PARAMETER NoHeader
            Does not put field names at the top of columns.
        .PARAMETER RangeName
            Makes the data in the worksheet a named range.
        .PARAMETER TableName
            Makes the data in the worksheet a table with a name applies a style to it. Name must not contain spaces.
        .PARAMETER TableStyle
            Selects the style for the named table - defaults to 'Medium6'.
        .PARAMETER ExcelChartDefinition
            A hash table containing ChartType, Title, NoLegend, ShowCategory, ShowPercent, Yrange, Xrange and SeriesHeader for one or more [non-pivot] charts.
        .PARAMETER HideSheet
            Name(s) of Sheet(s) to hide in the workbook.
        .PARAMETER MoveToStart 
            If specified, the worksheet will be moved to the start of the workbook.
            MoveToStart takes precedence over MoveToEnd, Movebefore and MoveAfter if more than one is specified.
        .PARAMETER MoveToEnd 
            If specified, the worksheet will be moved to the end of the workbook. 
            (This is the default position for newly created sheets, but this can be used to move existing sheets.) 
        .PARAMETER MoveBefore
            If specified, the worksheet will be moved before the nominated one (which can be a postion starting from 1, or a name). 
            MoveBefore takes precedence over MoveAfter if both are specified.
        .PARAMETER MoveAfter
            If specified, the worksheet will be moved after the nominated one (which can be a postion starting from 1, or a name or *).
            If * is used, the worksheet names will be examined starting with the first one, and the sheet placed after the last sheet which comes before it alphabetically.   
        .PARAMETER KillExcel
            Closes Excel - prevents errors writing to the file because Excel has it open.
        .PARAMETER AutoNameRange
            Makes each column a named range.
        .PARAMETER StartRow
            Row to start adding data. 1 by default. Row 1 will contain the title if any. Then headers will appear (Unless -No header is specified) then the data appears.
        .PARAMETER StartColumn
            Column to start adding data - 1 by default.
        .PARAMETER FreezeTopRow
            Freezes headers etc. in the top row.
        .PARAMETER FreezeFirstColumn
            Freezes titles etc. in the left column.
        .PARAMETER FreezeTopRowFirstColumn
             Freezes top row and left column (equivalent to Freeze pane 2,2 ).
        .PARAMETER FreezePane
             Freezes panes at specified coordinates (in the form  RowNumber , ColumnNumber).
        .PARAMETER AutoFilter
            Enables the 'Filter' in Excel on the complete header row. So users can easily sort, filter and/or search the data in the select column from within Excel.
        .PARAMETER AutoSize
            Sizes the width of the Excel column to the maximum width needed to display all the containing data in that cell.
        .PARAMETER Now
            The 'Now' switch is a shortcut that creates automatically a temporary file, enables 'AutoSize', 'AutoFiler' and 'Show', and opens the file immediately.
        .PARAMETER NumberFormat
            Formats all values that can be converted to a number to the format specified.

            Examples:
            # integer (not really needed unless you need to round numbers, Excel will use default cell properties).
            '0'

            # integer without displaying the number 0 in the cell.
            '#'

            # number with 1 decimal place.
            '0.0'

            # number with 2 decimal places.
            '0.00'

            # number with 2 decimal places and thousand separator.
            '#,##0.00'

            # number with 2 decimal places and thousand separator and money symbol.
            '€#,##0.00'

            # percentage (1 = 100%, 0.01 = 1%)
            '0%'

            # Blue color for positive numbers and a red color for negative numbers. All numbers will be proceeded by a dollar sign '$'.
            '[Blue]$#,##0.00;[Red]-$#,##0.00'

        .PARAMETER Show
            Opens the Excel file immediately after creation. Convenient for viewing the results instantly without having to search for the file first.
        .PARAMETER PassThru
            If specified, Export-Excel returns an object representing the Excel package without saving the package first. To save it you need to call the save or Saveas method or send it back to Export-Excel.

        .EXAMPLE
            Get-Process | Export-Excel .\Test.xlsx -show
            Export all the processes to the Excel file 'Test.xlsx' and open the file immediately.

        .EXAMPLE
            $ExcelParams = @{
                Path    = $env:TEMP + '\Excel.xlsx'
                Show    = $true
                Verbose = $true
            }
            Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
            Write-Output -1 668 34 777 860 -0.5 119 -0.1 234 788 |
                Export-Excel @ExcelParams -NumberFormat '[Blue]$#,##0.00;[Red]-$#,##0.00'

            Exports all data to the Excel file 'Excel.xslx' and colors the negative values in 'Red' and the positive values in 'Blue'. It will also add a dollar sign '$' in front of the rounded numbers to two decimal characters behind the comma.

        .EXAMPLE
            $ExcelParams = @{
                Path    = $env:TEMP + '\Excel.xlsx'
                Show    = $true
                Verbose = $true
            }
            Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
            [PSCustOmobject][Ordered]@{
                Date      = Get-Date
                Formula1  = '=SUM(F2:G2)'
                String1   = 'My String'
                String2   = 'a'
                IPAddress = '10.10.25.5'
                Number1   = '07670'
                Number2   = '0,26'
                Number3   = '1.555,83'
                Number4   = '1.2'
                Number5   = '-31'
                PhoneNr1  = '+32 44'
                PhoneNr2  = '+32 4 4444 444'
                PhoneNr3  =  '+3244444444'
            } | Export-Excel @ExcelParams -NoNumberConversion IPAddress, Number1

            Exports all data to the Excel file 'Excel.xslx' and tries to convert all values to numbers where possible except for 'IPAddress' and 'Number1'. These are stored in the sheet 'as is', without being converted to a number.

        .EXAMPLE
            $ExcelParams = @{
                Path    = $env:TEMP + '\Excel.xlsx'
                Show    = $true
                Verbose = $true
            }
            Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
            [PSCustOmobject][Ordered]@{
                Date      = Get-Date
                Formula1  = '=SUM(F2:G2)'
                String1   = 'My String'
                String2   = 'a'
                IPAddress = '10.10.25.5'
                Number1   = '07670'
                Number2   = '0,26'
                Number3   = '1.555,83'
                Number4   = '1.2'
                Number5   = '-31'
                PhoneNr1  = '+32 44'
                PhoneNr2  = '+32 4 4444 444'
                PhoneNr3  =  '+3244444444'
            } | Export-Excel @ExcelParams -NoNumberConversion *

            Exports all data to the Excel file 'Excel.xslx' as is, no number conversion will take place. This means that Excel will show the exact same data that you handed over to the 'Export-Excel' function.

        .EXAMPLE
            $ExcelParams = @{
                Path    = $env:TEMP + '\Excel.xlsx'
                Show    = $true
                Verbose = $true
            }
            Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
            Write-Output 489 668 299 777 860 151 119 497 234 788 |
                Export-Excel @ExcelParams -ConditionalText $(
                    New-ConditionalText -ConditionalType GreaterThan 525 -ConditionalTextColor DarkRed -BackgroundColor LightPink
                )

            Exports data that will have a 'Conditional formatting rule' in Excel on these cells that will show the background fill color in 'LightPink' and the text color in 'DarkRed' when the value is greater then '525'. In case this condition is not met the color will be the default, black text on a white background.

        .EXAMPLE
            $ExcelParams = @{
                Path    = $env:TEMP + '\Excel.xlsx'
                Show    = $true
                Verbose = $true
            }
            Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
            Get-Service | Select Name, Status, DisplayName, ServiceName |
                Export-Excel @ExcelParams -ConditionalText $(
                    New-ConditionalText Stop DarkRed LightPink
                    New-ConditionalText Running Blue Cyan
                )

            Export all services to an Excel sheet where all cells have a 'Conditional formatting rule' in Excel that will show the background fill color in 'LightPink' and the text color in 'DarkRed' when the value contains the word 'Stop'. If the value contains the word 'Running' it will have a background fill color in 'Cyan' and a text color 'Blue'. In case none of these conditions are met the color will be the default, black text on a white background.

        .EXAMPLE
            $ExcelParams = @{
                Path      = $env:TEMP + '\Excel.xlsx'
                Show      = $true
                Verbose   = $true
            }
            Remove-Item -Path $ExcelParams.Path -Force -EA Ignore

            $Array = @()

            $Obj1 = [PSCustomObject]@{
                Member1   = 'First'
                Member2   = 'Second'
            }

            $Obj2 = [PSCustomObject]@{
                Member1   = 'First'
                Member2   = 'Second'
                Member3   = 'Third'
            }

            $Obj3 = [PSCustomObject]@{
                Member1   = 'First'
                Member2   = 'Second'
                Member3   = 'Third'
                Member4   = 'Fourth'
            }

            $Array = $Obj1, $Obj2, $Obj3
            $Array | Out-GridView -Title 'Not showing Member3 and Member4'
            $Array | Update-FirstObjectProperties | Export-Excel @ExcelParams -WorkSheetname Numbers

            Updates the first object of the array by adding property 'Member3' and 'Member4'. Afterwards. all objects are exported to an Excel file and all column headers are visible.

        .EXAMPLE
            Get-Process | Export-Excel .\test.xlsx -WorkSheetname Processes -IncludePivotTable -Show -PivotRows Company -PivotData PM

        .EXAMPLE
            Get-Process | Export-Excel .\test.xlsx -WorkSheetname Processes -ChartType PieExploded3D -IncludePivotChart -IncludePivotTable -Show -PivotRows Company -PivotData PM

        .EXAMPLE
            Get-Service | Export-Excel 'c:\temp\test.xlsx'  -Show -IncludePivotTable -PivotRows status -PivotData @{status='count'}

        .EXAMPLE
            $pt = [ordered]@{}
            $pt.pt1=@{ SourceWorkSheet   = 'Sheet1';
                       PivotRows         = 'Status'
                       PivotData         = @{'Status'='count'}
                       IncludePivotChart = $true
                       ChartType         = 'BarClustered3D'
            }
            $pt.pt2=@{ SourceWorkSheet   = 'Sheet2';
                       PivotRows         = 'Company'
                       PivotData         = @{'Company'='count'}
                       IncludePivotChart = $true
                       ChartType         = 'PieExploded3D'
            }
            Remove-Item  -Path .\test.xlsx
            Get-Service | Select-Object    -Property Status,Name,DisplayName,StartType | Export-Excel -Path .\test.xlsx -AutoSize
            Get-Process | Select-Object    -Property Name,Company,Handles,CPU,VM       | Export-Excel -Path .\test.xlsx -AutoSize -WorkSheetname 'sheet2'
            Export-Excel -Path .\test.xlsx -PivotTableDefinition $pt -Show

            This example defines two pivot tables. Then it puts Service data on Sheet1 with one call to Export-Excel and Process Data on sheet2 with a second call to Export-Excel.
            The thrid and final call adds the two pivot tables and opens the spreadsheet in Excel.


        .EXAMPLE
            Remove-Item  -Path .\test.xlsx
            $excel = Get-Service | Select-Object -Property Status,Name,DisplayName,StartType | Export-Excel -Path .\test.xlsx -PassThru
            $excel.Workbook.Worksheets["Sheet1"].Row(1).style.font.bold = $true
            $excel.Workbook.Worksheets["Sheet1"].Column(3 ).width = 29
            $excel.Workbook.Worksheets["Sheet1"].Column(3 ).Style.wraptext = $true
            $excel.Save()
            $excel.Dispose()
            Start-Process .\test.xlsx

            This example uses -passthrough - put service information into sheet1 of the work book and saves the excelPackageObject in $Excel.
            It then uses the package object to apply formatting, and then saves the workbook and disposes of the object before loading the document in Excel.

        .EXAMPLE
            $excel = Get-Process | Select-Object -Property Name,Company,Handles,CPU,PM,NPM,WS | Export-Excel -Path .\test.xlsx -ClearSheet -WorkSheetname "Processes" -PassThru
            $sheet = $excel.Workbook.Worksheets["Processes"]
            $sheet.Column(1) | Set-Format -Bold -AutoFit
            $sheet.Column(2) | Set-Format -Width 29 -WrapText
            $sheet.Column(3) | Set-Format -HorizontalAlignment Right -NFormat "#,###"
            Set-Format -Address $sheet.Cells["E1:H1048576"]  -HorizontalAlignment Right -NFormat "#,###"
            Set-Format -Address $sheet.Column(4)  -HorizontalAlignment Right -NFormat "#,##0.0" -Bold
            Set-Format -Address $sheet.Row(1) -Bold -HorizontalAlignment Center
            Add-ConditionalFormatting -WorkSheet $sheet -Range "D2:D1048576" -DataBarColor Red
            Add-ConditionalFormatting -WorkSheet $sheet -Range "G2:G1048576" -RuleType GreaterThan -ConditionValue "104857600" -ForeGroundColor Red
            foreach ($c in 5..9) {Set-Format $sheet.Column($c)  -AutoFit }
            Export-Excel -ExcelPackage $excel -WorkSheetname "Processes" -IncludePivotChart -ChartType ColumnClustered -NoLegend -PivotRows company  -PivotData @{'Name'='Count'}  -Show

            This a more sophisticated version of the previous example showing different ways of using Set-Format, and also adding conditional formatting.
            In the final command a Pivot chart is added and the workbook is opened in Excel.

        .LINK
            https://github.com/dfinke/ImportExcel
    #>

    [CmdletBinding(DefaultParameterSetName = 'Default')]
    Param(
        [Parameter(ParameterSetName = "Default", Position = 0)]
        [Parameter(ParameterSetName = "Table"  , Position = 0)]
        [String]$Path,
        [Parameter(Mandatory = $true, ParameterSetName = "PackageDefault")]
        [Parameter(Mandatory = $true, ParameterSetName = "PackageTable")]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        [Parameter(ValueFromPipeline = $true)]
        $TargetData,
        [Switch]$Show,
        [String]$WorkSheetname = 'Sheet1',
        [String]$Password,
        [switch]$ClearSheet,
        [switch]$Append,
        [String]$Title,
        [OfficeOpenXml.Style.ExcelFillStyle]$TitleFillPattern = 'None',
        [Switch]$TitleBold,
        [Int]$TitleSize = 22,
        [System.Drawing.Color]$TitleBackgroundColor,
        [Switch]$IncludePivotTable,
        [String[]]$PivotRows,
        [String[]]$PivotColumns,
        $PivotData,
        [String[]]$PivotFilter,
        [Switch]$PivotDataToColumn,
        [Hashtable]$PivotTableDefinition,
        [Switch]$IncludePivotChart,
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = 'Pie',
        [Switch]$NoLegend,
        [Switch]$ShowCategory,
        [Switch]$ShowPercent,
        [Switch]$AutoSize,
        [Switch]$NoClobber,
        [Switch]$FreezeTopRow,
        [Switch]$FreezeFirstColumn,
        [Switch]$FreezeTopRowFirstColumn,
        [Int[]]$FreezePane,
        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'PackageDefault')]
        [Switch]$AutoFilter,
        [Switch]$BoldTopRow,
        [Switch]$NoHeader,
        [String]$RangeName,
        [ValidateScript( {
                if ($_.Contains(' ')) {
                    throw 'Tablename has spaces.'
                }
                elseif (-not $_) {
                    throw 'Tablename is null or empty.'
                }
                elseif ($_[0] -notmatch '[a-z]') {
                    throw 'Tablename starts with an invalid character.'
                }
                else {
                    $true
                }
            })]
        [Parameter(ParameterSetName = 'Table'        , Mandatory = $true)]
        [Parameter(ParameterSetName = 'PackageTable' , Mandatory = $true)]
        [String]$TableName,
        [Parameter(ParameterSetName = 'Table')]
        [Parameter(ParameterSetName = 'PackageTable')]
        [OfficeOpenXml.Table.TableStyles]$TableStyle = 'Medium6',
        [Object[]]$ExcelChartDefinition,
        [String[]]$HideSheet,
        [Switch]$MoveToStart,  
        [Switch]$MoveToEnd,  
        $MoveBefore ,
        $MoveAfter ,
        [Switch]$KillExcel,
        [Switch]$AutoNameRange,
        [Int]$StartRow = 1,
        [Int]$StartColumn = 1,
        [Switch]$PassThru,
        [String]$Numberformat = 'General',
        [string[]]$ExcludeProperty,
        [String[]]$NoNumberConversion,
        [Object[]]$ConditionalFormat,
        [Object[]]$ConditionalText,
        [ScriptBlock]$CellStyleSB,
        [Parameter(ParameterSetName = 'Now')]
        # [Parameter(ParameterSetName = 'TableNow')]
        [Switch]$Now,
        [Switch]$ReturnRange,
        [Switch]$NoTotalsInPivot,
        [Switch]$ReZip
    )

    Begin {
        function Add-CellValue {
            <#
              .SYNOPSIS
                Save a value in an Excel cell.

              .DESCRIPTION
                DateTime objects are always converted to a short DateTime format in Excel. When Excel loads the file,
                it applies the local format for dates. And formulas are always saved as formulas. URIs are set as hyperlinks in the file.

                Numerical values will be converted to numbers as defined in the regional settings of the local
                system. In case the parameter 'NoNumberConversion' is used, we don't convert to number and leave
                the value 'as is'. In case of conversion failure, we also leave the value 'as is'.
            #>

            Param (
                [Object]$TargetCell,
                [Object]$CellValue
            )

            Switch ($CellValue) {
                {($_ -is [String]) -and ($_.StartsWith('='))} {
                    #region Save an Excel formula
                    $TargetCell.Formula = $_
                    Write-Verbose "Cell '$Row`:$ColumnIndex' header '$Name' add value '$_' as formula"
                    break
                    #endregion
                }
                { $_ -is [URI] } {
                    #region Save a hyperlink
                    $TargetCell.Value = $_.AbsoluteUri
                    $TargetCell.HyperLink = $_ 
                    $TargetCell.Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
                    $TargetCell.Style.Font.UnderLine = $true
                    Write-Verbose "Cell '$Row`:$ColumnIndex' header '$Name' add value '$($_.AbsoluteUri)' as Hyperlink"
                    break
                    #endregion
                }
                { $_ -is [DateTime]} {
                    #region Save a date with an international valid format
                    $TargetCell.Value = $_
                    $TargetCell.Style.Numberformat.Format = 'm/d/yy h:mm' # This is not a custom format, but a preset recognized as date and localized.
                    Write-Verbose "Cell '$Row`:$ColumnIndex' header '$Name' add value '$_' as date"
                    break
                    #endregion
                }

                {(($NoNumberConversion) -and ($NoNumberConversion -contains $Name)) -or
                    ($NoNumberConversion -eq '*')} {
                    #region Save a value without converting to number
                    $TargetCell.Value = $_
                    Write-Verbose "Cell '$Row`:$ColumnIndex' header '$Name' add value '$($TargetCell.Value)' unconverted"
                    break
                    #endregion
                }

                Default {
                    #region Save a value as a number if possible                  
                    $number = $null 
                    if ([Double]::TryParse([String]$_, [System.Globalization.NumberStyles]::Any,
                            [System.Globalization.NumberFormatInfo]::CurrentInfo, [Ref]$number)) {
                        $TargetCell.Value = $number
                        $targetCell.Style.Numberformat.Format = $Numberformat
                        Write-Verbose "Cell '$Row`:$ColumnIndex' header '$Name' add value '$($TargetCell.Value)' as number converted from '$_' with format '$Numberformat'"
                    }
                    else {
                        $TargetCell.Value = $_
                        Write-Verbose "Cell '$Row`:$ColumnIndex' header '$Name' add value '$($TargetCell.Value)' as string"
                    }
                    break
                    #endregion
                }
            }
        }

        Try {
            $script:Header = $null
            if ($Append -and $ClearSheet) {throw "You can't use -Append AND -ClearSheet."}

            if ($PSBoundParameters.Keys.Count -eq 0 -Or $Now) {
                $Path = [System.IO.Path]::GetTempFileName() -replace '\.tmp', '.xlsx'
                $Show = $true
                $AutoSize = $true
                if (!$TableName) {
                    $AutoFilter = $true
                }
            }

            if ($ExcelPackage) {
                   $pkg  = $ExcelPackage
                   $Path = $pkg.File
            }
            Else { $pkg  = Open-ExcelPackage -Path $Path -Create -KillExcel:$KillExcel}
           
            $params =  @{}
            if ($NoClobber) {Write-Warning -Message "-NoClobber parameter is no longer used" }
            foreach ($p in @("WorkSheetname","ClearSheet","MoveToStart","MoveToEnd","MoveBefore","MoveAfter")) {if ($PSBoundParameters[$p]) {$params[$p] = $PSBoundParameters[$p]}}
            $ws = $pkg | Add-WorkSheet @params  

            $ws.View.TabSelected = $true
            foreach ($format in $ConditionalFormat ) {
                $target = "Add$($format.Formatter)"
                $rule = ($ws.ConditionalFormatting).PSObject.Methods[$target].Invoke($format.Range, $format.IconType)
                $rule.Reverse = $format.Reverse
            }

            if ($append -and $ws.Dimension) {
                $headerRange = $ws.Dimension.Address -replace "\d+$", "1"
                #if there is a title or anything else above the header row, specifying StartRow will skip it.
                if ($StartRow -ne 1) {$headerRange = $headerRange -replace "1", "$StartRow"}
                #$script:Header     = $ws.Cells[$headerrange].Value
                #using a slightly odd syntax otherwise header ends up as a 2D array
                $ws.Cells[$headerRange].Value | ForEach-Object -Begin {$Script:header = @()} -Process {$Script:header += $_ }
                $row = $ws.Dimension.Rows
                Write-Debug -Message ("Appending: headers are " + ($script:Header -join ", ") + "Start row $row")
            }
            elseif ($Title) {
                #Can only add a title if not appending!
                $Row = $StartRow
                $ws.Cells[$Row, $StartColumn].Value = $Title
                $ws.Cells[$Row, $StartColumn].Style.Font.Size = $TitleSize

                if ($TitleBold) {
                    #Set title to Bold face font if -TitleBold was specified.
                    #Otherwise the default will be unbolded.
                    $ws.Cells[$Row, $StartColumn].Style.Font.Bold = $True
                }
                #Can only set TitleBackgroundColor if TitleFillPattern is something other than None.
                if ($TitleBackgroundColor -and ($TitleFillPattern -eq 'None')) {
                    $TitleFillPattern = [OfficeOpenXml.Style.ExcelFillStyle]::Solid 
                } 
                $ws.Cells[$Row, $StartColumn].Style.Fill.PatternType = $TitleFillPattern

                if ($TitleBackgroundColor ) {
                    $ws.Cells[$Row, $StartColumn].Style.Fill.BackgroundColor.SetColor($TitleBackgroundColor)
                }
                $Row ++ ; $startRow ++
                }
            else {  $Row = $StartRow }
            $ColumnIndex = $StartColumn
            $firstTimeThru = $true
            $isDataTypeValueType = $false
        }
        Catch {
            if ($AlreadyExists) {
                #Is this set anywhere ?
                throw "Failed exporting worksheet '$WorkSheetname' to '$Path': The worksheet '$WorkSheetname' already exists."
            }
            else {
                throw "Failed exporting worksheet '$WorkSheetname' to '$Path': $_"
            }
        }
    }

    Process {
        if ($TargetData) {
            Try {
                if ($firstTimeThru) {
                    $firstTimeThru = $false
                    $isDataTypeValueType = $TargetData.GetType().name -match 'string|bool|byte|char|decimal|double|float|int|long|sbyte|short|uint|ulong|ushort' 
                    Write-Debug "DataTypeName is '$($TargetData.GetType().name)' isDataTypeValueType '$isDataTypeValueType'"
                }

                if ($isDataTypeValueType) {
                    $ColumnIndex = $StartColumn

                    Add-CellValue -TargetCell $ws.Cells[$Row, $ColumnIndex] -CellValue $TargetData

                    $Row += 1
                }
                else {
                    #region Add headers
                    if (-not $script:Header) {
                        $ColumnIndex = $StartColumn
                        $script:Header = $TargetData.PSObject.Properties.Name | Where-Object {$_ -notin $ExcludeProperty}

                        if ($NoHeader) {
                            # Don't push the headers to the spreadsheet
                            $Row -= 1
                        }
                        else {
                            foreach ($Name in $script:Header) {
                                $ws.Cells[$Row, $ColumnIndex].Value = $Name
                                Write-Verbose "Cell '$Row`:$ColumnIndex' add header '$Name'"
                                $ColumnIndex += 1
                            }
                        }
                    }
                    #endregion

                    $Row += 1
                    $ColumnIndex = $StartColumn

                    foreach ($Name in $script:Header) {
                        #region Add non header values
                        Add-CellValue -TargetCell $ws.Cells[$Row, $ColumnIndex] -CellValue $TargetData.$Name

                        $ColumnIndex += 1
                        #endregion
                    }
                }
            }
            Catch {
                throw "Failed exporting worksheet '$WorkSheetname' to '$Path': $_"
            }
        }
    }

    End {
        Try {
            if ($AutoNameRange) {
                if (-not $script:header) {
                    $headerRange = $ws.Dimension.Address -replace "\d+$", "1"
                    #if there is a title or anything else above the header row, specifying StartRow will skip it.
                    if ($StartRow -ne 1) {$headerRange = $headerRange -replace "1", "$StartRow"}
                    #using a slightly odd syntax otherwise header ends up as a 2D array
                    $ws.Cells[$headerRange].Value | ForEach-Object -Begin {$Script:header = @()} -Process {$Script:header += $_ }
                }
                $totalRows = $ws.Dimension.End.Row
                $totalColumns = $ws.Dimension.Columns
                foreach ($c in 0..($totalColumns - 1)) {
                    $targetRangeName = "$($script:Header[$c])"
                    $targetColumn = $c + $StartColumn
                    $theCell = $ws.Cells[($startrow + 1), $targetColumn, $totalRows , $targetColumn ]
                    if ($ws.names[$targetRangeName]) { $ws.names[$targetRangeName].Address = $theCell.FullAddressAbsolute }
                    else {$ws.Names.Add($targetRangeName, $theCell) | Out-Null }

                    if ([OfficeOpenXml.FormulaParsing.ExcelUtilities.ExcelAddressUtil]::IsValidAddress($targetRangeName)) {
                        Write-Warning "AutoNameRange: Property name '$targetRangeName' is also a valid Excel address and may cause issues. Consider renaming the property name."
                    }
                }
            }

            if ($Title) {
                $startAddress = $ws.Dimension.Start.address -replace "$($ws.Dimension.Start.row)`$", "$($ws.Dimension.Start.row + 1)"
            }
            else {
                $startAddress = $ws.Dimension.Start.Address
            }

            $dataRange = "{0}:{1}" -f $startAddress, $ws.Dimension.End.Address

            Write-Debug "Data Range '$dataRange'"

            if (-not [String]::IsNullOrEmpty($RangeName)) {
                if ($ws.Names[$RangeName]) { $ws.Names[$rangename].Address = $ws.Cells[$dataRange].FullAddressAbsolute }
                else {$ws.Names.Add($RangeName, $ws.Cells[$dataRange]) | Out-Null } 
            }

            if (-not [String]::IsNullOrEmpty($TableName)) {
                $csr = $StartRow

                $csc = $StartColumn
                $cer = $ws.Dimension.End.Row
                $cec = $ws.Dimension.End.Column # was $script:Header.Count

                $targetRange = $ws.Cells[$csr, $csc, $cer, $cec]
                #if we're appending data the table may already exist.  
                if ($ws.Tables[$TableName]) {
                    $ws.Tables[$TableName].TableXml.table.ref = $targetRange.Address 
                    $ws.Tables[$TableName].TableStyle = $TableStyle
                }
                else {
                    $tbl = $ws.Tables.Add($targetRange, $TableName)
                    $tbl.TableStyle = $TableStyle
                }
            }
            
            if ($PivotTableDefinition) {
                foreach ($item in $PivotTableDefinition.GetEnumerator()) {
                    $pivotTableName = $item.Key
                    $pivotTableDataName = $item.Key + 'PivotTableData'
                    if ($item.Value.PivotFilter) {$PivotTableStartCell = "A3"} else { $PivotTableStartCell = "A1"} 
                   
                    #Make sure the Pivot table sheet doesn't already exist.
                    #try {      $pkg.Workbook.Worksheets.Delete(    $pivotTableName) } catch {}
                    [OfficeOpenXml.ExcelWorksheet]$wsPivot = $pkg | Add-WorkSheet -WorkSheetname $pivotTableName -NoClobber:$NoClobber
                    
                    #If it is a pivot for the default sheet and it doesn't exist - create it  
                    if (-not $item.Value.SourceWorkSheet -and -not $wsPivot.PivotTables[$pivotTableDataName] ) { 
                        $pivotTable = $wsPivot.PivotTables.Add($wsPivot.Cells[$PivotTableStartCell], $ws.Cells[$dataRange], $pivotTableDataName)
                    } 
                    #If it is a pivot for the default sheet and it exists - update the range. 
                    elseif (-not $item.Value.SourceWorkSheet -and $wsPivot.PivotTables[$pivotTableDataName] ) { 
                        $wsPivot.PivotTables[$pivotTableDataName].CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.ref = $WS.Cells[$dataRange].Address
                    } 
                    #if it is a pivot for a named sheet and it doesn't exist, create it. 
                    elseif ($item.Value.SourceWorkSheet -and -not $wsPivot.PivotTables[$pivotTableDataName] ) {
                        #find the worksheet
                        $workSheet = $pkg.Workbook.Worksheets.where( {$_.name -match $item.Value.SourceWorkSheet})[0]  
                        if (-not $workSheet) {Write-Warning -Message "Could not find Worksheet '$($item.Value.SourceWorkSheet)' specified in pivot-table definition $($item.key)." }
                        else {
                            if ($item.Value.SourceRange) { $targetdataRange = $item.Value.SourceRange } 
                            else { $targetDataRange =  $workSheet.Dimension.Address} 
                            $pivotTable = $wsPivot.PivotTables.Add($wsPivot.Cells[$PivotTableStartCell], $workSheet.Cells[$targetDataRange], $pivotTableDataName)
                        } 
                    }

                    #if we created the pivot table, set up the rows, columns and data if we didn't, put out a message 'existed' or 'error' .
                    if ($pivotTable) { 
                        foreach ($Row in $item.Value.PivotRows) {
                            try {$null = $pivotTable.RowFields.Add($pivotTable.Fields[$Row]) }
                            catch {Write-Warning -message "Could not add '$row' to Rows in PivotTable $pivotTableName." } 
                        }
                        foreach ($Column in $item.Value.PivotColumns) {
                            try {$null = $pivotTable.ColumnFields.Add($pivotTable.Fields[$Column])}
                            catch {Write-Warning -message "Could not add '$Column' to Columns in PivotTable $pivotTableName." } 
                        } 
                        if ($item.Value.PivotData -is [HashTable] -or $item.Value.PivotData -is [System.Collections.Specialized.OrderedDictionary]) {
                            $item.Value.PivotData.Keys | ForEach-Object {
                                try {
                                    $df = $pivotTable.DataFields.Add($pivotTable.Fields[$_])
                                    $df.Function = $item.Value.PivotData.$_ 
                                }
                                catch {Write-Warning -message "Problem adding data fields to PivotTable $pivotTableName." } 
                            }
                        }
                        else {
                            foreach ($field in $item.Value.PivotData) {
                                try {
                                    $df = $pivotTable.DataFields.Add($pivotTable.Fields[$field])
                                    $df.Function = 'Count'
                                }
                                catch {Write-Warning -message "Problem adding data field '$field' to PivotTable $pivotTableName." } 
                            }
                        }
                        foreach ( $pFilter in $item.Value.PivotFilter) {
                            try { $null = $pivotTable.PageFields.Add($pivotTable.Fields[$pFilter])}
                            catch {Write-Warning -message "Could not add '$pFilter' to Filter/Page fields in PivotTable $pivotTableName." } 
                        }
                        if ($item.Value.NoTotalsInPivot -or $NoTotalsInPivot) { $pivotTable.RowGrandTotals = $false }                        
                        if ($item.Value.PivotDataToColumn -or $PivotDataToColumn) { $pivotTable.DataOnRows = $false }                         
                    }
                    elseif ($wsPivot.PivotTables[$pivotTableDataName]) {
                        Write-Warning -Message "Pivot table defined in $($item.key) already exists."
                    } 
                    else {  Write-Warning -Message "Could not create the pivot table defined in $($item.key)."}
                    
                    #Create the chart if it doesn't exist, leave alone if it does. 
                    if ($item.Value.IncludePivotChart -and -not $wsPivot.Drawings['PivotChart'] ) {
                        if ($item.Value.ChartType) { $ChartType = $item.Value.ChartType} # $ChartType may be passed as a parameter, has default of "Pie", over-ride that if it is in the pivot definition
                        [OfficeOpenXml.Drawing.Chart.ExcelChart] $chart = $wsPivot.Drawings.AddChart('PivotChart', $ChartType, $pivotTable)
                        if (-not $item.Value.ChartHeight)                {$item.Value.ChartHeight = 400 }
                        if (-not $item.Value.ChartWidth)                 {$item.Value.ChartWidth  = 600 }
                        if (-not $item.Value.ChartRow)                   {$item.Value.ChartRow    = 0   }
                        if (-not $item.Value.ChartColumn)                {$item.Value.ChartColumn = 4   }
                        if (-not $item.Value.ChartRowOffSetPixels)       {$item.Value.ChartRowOffSetPixels     = 0 }
                        if (-not $item.Value.ChartColumnOffSetPixels)    {$item.Value.ChartColumnOffSetPixels  = 0 }
                        $chart.SetPosition($item.Value.ChartRow  ,        $item.Value.ChartRowOffSetPixels , $item.Value.ChartColumn, $item.Value.ChartColumnOffSetPixels)  
                        $chart.SetSize(    $item.Value.ChartWidth,        $item.Value.ChartHeight)
                        if ($chart.DataLabel) {
                            $chart.DataLabel.ShowCategory      = [boolean]$item.Value.ShowCategory
                            $chart.DataLabel.ShowPercent       = [boolean]$item.Value.ShowPercent
                        }
                        if ([boolean]$item.Value.NoLegend -or $NoLegend) {$chart.Legend.Remove()}
                        if (         $item.Value.ChartTitle)             {$chart.Title.Text  = $item.Value.chartTitle}
                    }
                }
            }

            if ($IncludePivotTable -or $IncludePivotChart) {
                if ($PivotFilter) {$PivotTableStartCell = "A3"} else {$PivotTableStartCell = "A1"} 

                $pivotTableName = $WorkSheetname + 'PivotTable'
                $wsPivot = $pkg | Add-WorkSheet -WorkSheetname $pivotTableName -NoClobber:$NoClobber

                $wsPivot.View.TabSelected = $true

                $pivotTableDataName = $WorkSheetname + 'PivotTableData'
                if ($wsPivot.PivotTables[$pivotTableDataName] ) {
                    $pivotTable = $wsPivot.PivotTables[$pivotTableDataName]
                    $pivotTable.CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.ref = $WS.Cells[$dataRange].Address
                    Write-Warning -Message "Pivot table for $worksheetName already exists; updating the data range, but other properties will not be changed" 
                }    
                else {
                    $pivotTable = $wsPivot.PivotTables.Add($wsPivot.Cells[$PivotTableStartCell], $ws.Cells[$dataRange], $pivotTableDataName)

                    foreach ($Row in $PivotRows) {
                        try {$null = $pivotTable.RowFields.Add($pivotTable.Fields[$Row]) }
                        catch {Write-Warning -message "Could not add '$row' to PivotTable Rows." } 
                    }
                    
                    foreach ($Column in $PivotColumns) {
                        try {$null = $pivotTable.ColumnFields.Add($pivotTable.Fields[$Column])}
                        catch {Write-Warning -message "Could not add '$Column' to PivotTable Columns." } 
                    }
                     
                    if ($PivotData -is [HashTable] -or $PivotData -is [System.Collections.Specialized.OrderedDictionary]) {
                        $PivotData.Keys | ForEach-Object {
                            try {
                                $df = $pivotTable.DataFields.Add($pivotTable.Fields[$_])
                                $df.Function = $PivotData.$_
                            }
                            catch {Write-Warning "Problem adding to Pivot table data fields." }
                        }
                    }
                    else {
                        foreach ($Item in $PivotData) {
                            try {
                                $df = $pivotTable.DataFields.Add($pivotTable.Fields[$Item])
                                $df.Function = 'Count'
                            }
                            catch {Write-Warning "Problem adding '$item' to Pivot table data fields." }
                        }
                    }
                    
                    if ($PivotDataToColumn) { $pivotTable.DataOnRows = $false }

                    foreach ($pFilter in $PivotFilter) {
                        try {$null = $pivotTable.PageFields.Add($pivotTable.Fields[$pFilter])  }
                        catch {Write-Warning "Problem adding 'pFilter' to Pivot table page/filter fields." }
                    }
                  
                    if ($NoTotalsInPivot) { $pivotTable.RowGrandTotals = $false }
                }

                if ($IncludePivotChart) {
                    if (-not $wsPivot.Drawings['PivotChart']) {
                        $chart = $wsPivot.Drawings.AddChart('PivotChart', $ChartType, $pivotTable) 
                        if ($chart.DataLabel) {
                            $chart.DataLabel.ShowCategory = $ShowCategory
                            $chart.DataLabel.ShowPercent = $ShowPercent
                        }
                        $chart.SetPosition(0, 26, 2, 26)  # if Pivot table is rows+data only it will be 2 columns wide if has pivot columns we don't know how wide it will be
                        if ($NoLegend) { $chart.Legend.Remove() }
                    } 
                }
            }

            if ($Password) {
                $ws.Protection.SetPassword($Password)
            }

            if ($AutoFilter) {
                $ws.Cells[$dataRange].AutoFilter = $true
            }

            if ($FreezeTopRow) {
                $ws.View.FreezePanes(2, 1)
            }

            if ($FreezeTopRowFirstColumn) {
                $ws.View.FreezePanes(2, 2)
            }

            if ($FreezeFirstColumn) {
                $ws.View.FreezePanes(1, 2)
            }

            if ($FreezePane) {
                $freezeRow, $freezeColumn = $FreezePane
                if (-not $freezeColumn -or $freezeColumn -eq 0) {
                    $freezeColumn = 1
                }

                if ($freezeRow -gt 1) {
                    $ws.View.FreezePanes($freezeRow, $freezeColumn)
                }
            }

            if ($BoldTopRow) {
                if ($Title) {
                    $range = $ws.Dimension.Address -replace '\d+', '2'
                }
                else {
                    $range = $ws.Dimension.Address -replace '\d+', '1'
                }

                $ws.Cells[$range].Style.Font.Bold = $true
            }

            if ($AutoSize) {
                $ws.Cells.AutoFitColumns()
            }

            foreach ($Sheet in $HideSheet) {
                $pkg.Workbook.WorkSheets[$Sheet].Hidden = 'Hidden'
            }

            foreach ($chartDef in $ExcelChartDefinition) {
                $ChartName = 'Chart' + (Split-Path -Leaf ([System.IO.path]::GetTempFileName())) -replace 'tmp|\.', ''
                $chart = $ws.Drawings.AddChart($ChartName, $chartDef.ChartType)
                $chart.Title.Text = $chartDef.Title

                if ($chartDef.NoLegend) {
                    $chart.Legend.Remove()
                }

                if ($chart.Datalabel -ne $null) {
                    $chart.Datalabel.ShowCategory = $chartDef.ShowCategory
                    $chart.Datalabel.ShowPercent = $chartDef.ShowPercent
                }

                $chart.SetPosition($chartDef.Row, $chartDef.RowOffsetPixels, $chartDef.Column, $chartDef.ColumnOffsetPixels)
                $chart.SetSize($chartDef.Width, $chartDef.Height)

                $chartDefCount = @($chartDef.YRange).Count
                if ($chartDefCount -eq 1) {
                    $Series = $chart.Series.Add($chartDef.YRange, $chartDef.XRange)

                    $SeriesHeader = $chartDef.SeriesHeader
                    if (-not $SeriesHeader) {
                        $SeriesHeader = 'Series 1'
                    }

                    $Series.Header = $SeriesHeader
                }
                else {
                    for ($idx = 0; $idx -lt $chartDefCount; $idx += 1) {
                        $Series = $chart.Series.Add($chartDef.YRange[$idx], $chartDef.XRange)

                        if ($chartDef.SeriesHeader.Count -gt 0) {
                            $SeriesHeader = $chartDef.SeriesHeader[$idx]
                        }

                        if (-not $SeriesHeader) {
                            $SeriesHeader = "Series $($idx)"
                        }

                        $Series.Header = $SeriesHeader
                        $SeriesHeader = $null
                    }
                }
            }

            if ($ConditionalText) {
                foreach ($targetConditionalText in $ConditionalText) {
                    $target = "Add$($targetConditionalText.ConditionalType)"

                    $Range = $targetConditionalText.Range
                    if (-not $Range) {
                        $Range = $ws.Dimension.Address
                    }

                    $rule = ($ws.Cells[$Range].ConditionalFormatting).PSObject.Methods[$target].Invoke()

                    if ($targetConditionalText.Text) {
                        if ($targetConditionalText.ConditionalType -match 'equal|notequal|lessthan|lessthanorequal|greaterthan|greaterthanorequal') {
                            $rule.Formula = $targetConditionalText.Text
                        }
                        else {
                            $rule.Text = $targetConditionalText.Text
                        }
                    }

                    $rule.Style.Font.Color.Color = $targetConditionalText.ConditionalTextColor
                    $rule.Style.Fill.PatternType = $targetConditionalText.PatternType
                    $rule.Style.Fill.BackgroundColor.Color = $targetConditionalText.BackgroundColor
                }
            }

            if ($CellStyleSB) {
                $TotalRows = $ws.Dimension.Rows
                $LastColumn = (Get-ExcelColumnName $ws.Dimension.Columns).ColumnName
                & $CellStyleSB $ws $TotalRows $LastColumn
            }

            if ($PassThru) {
                $pkg
            }
            else {
                if ($ReturnRange) {
                    $ws.Dimension.Address
                }

                $pkg.Save()

                if ($ReZip) {
                    write-verbose "Re-Zipping $($pkg.file) using .NET ZIP library"
                    $zipAssembly = "System.IO.Compression.Filesystem"
                    try {
                        Add-Type -assembly $zipAssembly -ErrorAction stop
                    } 
                    catch {
                        write-error "The -ReZip parameter requires .NET Framework 4.5 or later to be installed. Recommend to install Powershell v4+"
                        continue
                    }
                    
                    $TempZipPath = Join-Path -path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
                    [io.compression.zipfile]::ExtractToDirectory($pkg.File, $TempZipPath) | Out-Null
                    Remove-Item $pkg.File -Force
                    [io.compression.zipfile]::CreateFromDirectory($TempZipPath, $pkg.File) | Out-Null
                }

                $pkg.Dispose()

                if ($Show) {
                    Invoke-Item $Path
                }
            }
        }
        Catch {
            throw "Failed exporting worksheet '$WorkSheetname' to '$Path': $_"
        }
    }
}

function New-PivotTableDefinition {
<#
  .Synopsis
    Creates Pivot table definitons for export excel 
  .Description
    Export-Excel allows a single Pivot table to be defined using the parameters -IncludePivotTable, -PivotColumns -PivotRows, 
      =PivotData, -PivotFilter, -NoTotalsInPivot, -PivotDataToColumn, -IncludePivotChart and -ChartType.
    Its -PivotTableDefintion paramater allows multiple pivot tables to be defined, with additional parameters.
    New-PivotTableDefinition is a convenient way to build these definitions. 
   .Example 
    $pt  = New-PivotTableDefinition -PivotTableName "PT1" -SourceWorkSheet "Sheet1" -PivotRows "Status"  -PivotData @{Status='Count' } -PivotFilter 'StartType' -IncludePivotChart  -ChartType BarClustered3D 
    $Pt += New-PivotTableDefinition -PivotTableName "PT2" -SourceWorkSheet "Sheet2" -PivotRows "Company" -PivotData @{Company='Count'} -IncludePivotChart  -ChartType PieExploded3D  -ShowPercent -ChartTitle "Breakdown of processes by company"
    Get-Service | Select-Object    -Property Status,Name,DisplayName,StartType | Export-Excel -Path .\test.xlsx -AutoSize
    Get-Process | Select-Object    -Property Name,Company,Handles,CPU,VM       | Export-Excel -Path .\test.xlsx -AutoSize -WorkSheetname 'sheet2'
    $excel = Export-Excel -Path .\test.xlsx -PivotTableDefinition $pt -Show 

    This is a re-work of one of the examples in Export-Excel - instead of writing out the pivot definition hash table it is built by calling New-PivotTableDefinition. 
#>
    param(
        [Parameter(Mandatory)]
        [Alias("PivtoTableName")]#Previous typo - use alias to avoid breaking scripts
        $PivotTableName,
        #Worksheet where the data is found 
        $SourceWorkSheet,
        #Address range in the worksheet e.g "A10:F20" - the first row must be column names: if not specified the whole sheet will be used/ 
        $SourceRange, 
        #Fields to set as rows in the Pivot table 
        $PivotRows,
        #A hash table in form "FieldName"="Function", where function is one of 
        #Average, Count, CountNums, Max, Min, Product, None, StdDev, StdDevP, Sum, Var, VarP
        [hashtable]$PivotData,
        #Fields to set as columns in the Pivot table 
        $PivotColumns,
        #Fields to use to filter in the Pivot table 
        $PivotFilter,
        [Switch]$PivotDataToColumn,
        [Switch]$NoTotalsInPivot,
        #If specified a chart Will be included. 
        [Switch]$IncludePivotChart,
        #Optional title for the pivot chart, by default the title omitted.
        [String]$ChartTitle,
        #Height of the chart in Pixels (400 by default)
        [int]$ChartHeight = 400 ,
        #Width of the chart in Pixels (600 by default)
        [int]$ChartWidth = 600,
        #Cell position of the top left corner of the chart, there will be this number of rows above the top edge of the chart (default is 0, chart starts at top edge of row 1). 
        [Int]$ChartRow = 0 ,
        #Cell position of the top left corner of the chart, there will be this number of cells to the left of the chart (default is 4, chart starts at left edge of column E) 
        [Int]$ChartColumn = 4,
        #Vertical offset of the chart from the cell corner.
        [Int]$ChartRowOffSetPixels = 0 ,
        #Horizontal offset of the chart from the cell corner.
        [Int]$ChartColumnOffSetPixels = 0,
        #Type of chart
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = 'Pie',
        #If specified hides the chart legend
        [Switch]$NoLegend,
        #if specified attaches the category to slices in a pie chart : not supported on all chart types, this may give errors if applied to an unsupported type. 
        [Switch]$ShowCategory,
        #If specified attaches percentages to slices in a pie chart.
        [Switch]$ShowPercent
    )
    $validDataFuntions = [system.enum]::GetNames([OfficeOpenXml.Table.PivotTable.DataFieldFunctions])

    if ($PivotData.values.Where({$_ -notin $validDataFuntions}) ) {
        Write-Warning -Message ("Pivot data functions might not be valid, they should be one of " + ($validDataFuntions -join ", ") + ".")
    }  

    $parameters = @{} + $PSBoundParameters
    $parameters.Remove('PivotTableName')

    @{$PivotTableName = $parameters}
}
