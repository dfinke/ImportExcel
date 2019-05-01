function Export-Excel {
    <#
        .SYNOPSIS
            Exports data to an Excel worksheet.
        .DESCRIPTION
            Exports data to an Excel file and where possible tries to convert numbers
            in text fields so Excel recognizes them as numbers instead of text.
             After all: Excel is a spreadsheet program used for number manipulation
             and calculations. If number conversion is not desired, use the
             parameter -NoNumberConversion *.
        .PARAMETER Path
            Path to a new or existing .XLSX file.
        .PARAMETER  ExcelPackage
            An object representing an Excel Package - usually this is returned by specifying -PassThru allowing multiple commands to work on the same workbook without saving and reloading each time.
        .PARAMETER WorksheetName
            The name of a sheet within the workbook - "Sheet1" by default.
        .PARAMETER ClearSheet
            If specified Export-Excel will remove any existing worksheet with the selected name. The Default behaviour is to overwrite cells in this sheet as needed (but leaving non-overwritten ones in place).
        .PARAMETER Append
            If specified data will be added to the end of an existing sheet, using the same column headings.
        .PARAMETER TargetData
            Data to insert onto the worksheet - this is usually provided from the pipeline.
        .PARAMETER DisplayPropertySet
            Many (but not all) objects have a hidden property named psStandardmembers with a child property DefaultDisplayPropertySet ; this parameter reduces the properties exported to those in this set.
        .PARAMETER NoAliasOrScriptPropeties
            Some objects duplicate existing properties by adding aliases, or have Script properties which take a long time to return a value and slow the export down, if specified this removes these properties
        .PARAMETER ExcludeProperty
            Specifies properties which may exist in the target data but should not be placed on the worksheet.
        .PARAMETER Calculate
            If specified a recalculation of the worksheet will be requested before saving.
        .PARAMETER Title
            Text of a title to be placed in the top left cell.
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
            Adds a PivotTable using the data in the worksheet.
        .PARAMETER PivotTableName
            If a PivotTable is created from command line parameters, specifies the name of the new sheet holding the pivot. Defaults to "WorksheetName-PivotTable".
        .PARAMETER PivotRows
            Name(s) of column(s) from the spreadsheet which will provide the Row name(s) in a PivotTable created from command line parameters.
        .PARAMETER PivotColumns
            Name(s) of columns from the spreadsheet which will provide the Column name(s) in a PivotTable created from command line parameters.
        .PARAMETER PivotFilter
            Name(s) columns from the spreadsheet which will provide the Filter name(s) in a PivotTable created from command line parameters.
        .PARAMETER PivotData
            In a PivotTable created from command line parameters, the fields to use in the table body are given as a Hash table in the form ColumnName = Average|Count|CountNums|Max|Min|Product|None|StdDev|StdDevP|Sum|Var|VarP.
        .PARAMETER PivotDataToColumn
            If there are multiple datasets in a PivotTable, by default they are shown as separate rows under the given row heading; this switch makes them separate columns.
        .PARAMETER NoTotalsInPivot
            In a PivotTable created from command line parameters, prevents the addition of totals to rows and columns.
        .PARAMETER PivotTotals
            By default, PivotTables have totals for each row (on the right) and for each column at the bottom. This allows just one or neither to be selected.
        .PARAMETER PivotTableDefinition
            Instead of describing a single PivotTable with multiple command-line parameters; you can use a HashTable in the form PivotTableName = Definition;
            Definition is itself a Hashtable with Sheet, PivotRows, PivotColumns, PivotData, IncludePivotChart and ChartType values.
        .PARAMETER IncludePivotChart
            Include a chart with the PivotTable - implies -IncludePivotTable.
        .PARAMETER ChartType
            The type for PivotChart (one of Excel's defined chart types).
        .PARAMETER NoLegend
            Exclude the legend from the PivotChart.
        .PARAMETER ShowCategory
            Add category labels to the PivotChart.
        .PARAMETER ShowPercent
            Add percentage labels to the PivotChart.
        .PARAMETER ConditionalFormat
            One or more conditional formatting rules defined with New-ConditionalFormattingIconSet.
        .PARAMETER ConditionalText
            Applies a Conditional formatting rule defined with New-ConditionalText. When specific conditions are met the format is applied.
        .PARAMETER NoNumberConversion
            By default we convert all values to numbers if possible, but this isn't always desirable. NoNumberConversion allows you to add exceptions for the conversion. Wildcards (like '*') are allowed.
        .PARAMETER BoldTopRow
            Makes the top row boldface.
        .PARAMETER NoHeader
            Does not put field names at the top of columns.
        .PARAMETER RangeName
            Makes the data in the worksheet a named range.
        .PARAMETER TableName
            Makes the data in the worksheet a table with a name, and applies a style to it. The name must not contain spaces. If a style is specified without a name, table1, table2 etc. will be used.
        .PARAMETER TableStyle
            Selects the style for the named table - if a name is specified without a style, 'Medium6' is used as a default.
        .PARAMETER BarChart
            Creates a "quick" bar chart using the first text column as labels and the first numeric column as values
        .PARAMETER ColumnChart
            Creates a "quick" column chart using the first text column as labels and the first numeric column as values
        .PARAMETER LineChart
            Creates a "quick" line chart using the first text column as labels and the first numeric column as values
        .PARAMETER PieChart
            Creates a "quick" pie chart using the first text column as labels and the first numeric column as values
        .PARAMETER ExcelChartDefinition
            A hash table containing ChartType, Title, NoLegend, ShowCategory, ShowPercent, Yrange, Xrange and SeriesHeader for one or more [non-Pivot] charts.
        .PARAMETER HideSheet
            Name(s) of Sheet(s) to hide in the workbook, supports wildcards. If the selection would cause all sheets to be hidden, the sheet being worked on will be revealed.
        .PARAMETER UnHideSheet
            Name(s) of Sheet(s) to reveal in the workbook, supports wildcards.
        .PARAMETER MoveToStart
            If specified, the worksheet will be moved to the start of the workbook.
            -MoveToStart takes precedence over -MoveToEnd, -Movebefore and -MoveAfter if more than one is specified.
        .PARAMETER MoveToEnd
            If specified, the worksheet will be moved to the end of the workbook.
            (This is the default position for newly created sheets, but this can be used to move existing sheets.)
        .PARAMETER MoveBefore
            If specified, the worksheet will be moved before the nominated one (which can be a position starting from 1, or a name).
            -MoveBefore takes precedence over -MoveAfter if both are specified.
        .PARAMETER MoveAfter
            If specified, the worksheet will be moved after the nominated one (which can be a position starting from 1, or a name or *).
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
             Freezes panes at specified coordinates (in the form  RowNumber, ColumnNumber).
        .PARAMETER AutoFilter
            Enables the Excel filter on the complete header row, so users can easily sort, filter and/or search the data in the selected column.
        .PARAMETER AutoSize
            Sizes the width of the Excel column to the maximum width needed to display all the containing data in that cell.
        .PARAMETER MaxAutoSizeRows
            Autosizing can be time consuming, so this sets a maximum number of rows to look at for the Autosize operation. Default is 1000; If 0 is specified ALL rows will be checked
        .PARAMETER Activate
            If there is already content in the workbook, a new sheet will not be active UNLESS Activate is specified; if a PivotTable is included it will be the active sheet
        .PARAMETER Now
            The -Now switch is a shortcut that automatically creates a temporary file, enables "AutoSize", "AutoFiler" and "Show", and opens the file immediately.
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

            # number with 2 decimal places and thousand-separator.
            '#,##0.00'

            # number with 2 decimal places and thousand-separator and money-symbol.
            '€#,##0.00'

            # percentage (1 = 100%, 0.01 = 1%)
            '0%'

            # Blue color for positive numbers and a red color for negative numbers. All numbers will be proceeded by a dollar sign '$'.
            '[Blue]$#,##0.00;[Red]-$#,##0.00'

        .PARAMETER ReZip
            If specified, Export-Excel will expand the contents of the .XLSX file (which is multiple files in a zip archive) and rebuild it.
        .PARAMETER NoClobber
            Not used. Left in to avoid problems with older scripts, it may be removed in future versions.
        .PARAMETER CellStyleSB
            A script block which is run at the end of the export to apply styles to cells (although it can be used for other purposes).
            The script block is given three paramaters; an object containing the current worksheet, the Total number of Rows and the number of the last column.
        .PARAMETER Show
            Opens the Excel file immediately after creation. Convenient for viewing the results instantly without having to search for the file first.
        .PARAMETER ReturnRange
            If specified, Export-Excel returns the range of added cells in the format "A1:Z100".
        .PARAMETER PassThru
            If specified, Export-Excel returns an object representing the Excel package without saving the package first.
            To save, you need to call Close-ExcelPackage or send the object back to Export-Excel, or use its .Save() or SaveAs() method.
        .EXAMPLE
            Get-Process | Export-Excel .\Test.xlsx -show
            Export all the processes to the Excel file 'Test.xlsx' and open the file immediately.

        .EXAMPLE
        >
        PS> $ExcelParams = @{
                Path    = $env:TEMP + '\Excel.xlsx'
                Show    = $true
                Verbose = $true
            }
            Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
            Write-Output -1 668 34 777 860 -0.5 119 -0.1 234 788 |
                Export-Excel @ExcelParams -NumberFormat '[Blue]$#,##0.00;[Red]-$#,##0.00'

            Exports all data to the Excel file 'Excel.xslx' and colors the negative values
            in Red and the positive values in Blue. It will also add a dollar sign in front
            of the numbers which use a thousand seperator and display to two decimal places.

        .EXAMPLE
        >
        PS> $ExcelParams = @{
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

            Exports all data to the Excel file "Excel.xlsx" and tries to convert all values
            to numbers where possible except for "IPAddress" and "Number1", which are
            stored in the sheet 'as is', without being converted to a number.

        .EXAMPLE
        >
        PS> $ExcelParams = @{
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

            Exports all data to the Excel file 'Excel.xslx' as is, no number conversion
            will take place. This means that Excel will show the exact same data that
            you handed over to the 'Export-Excel' function.

        .EXAMPLE
        >
        PS> $ExcelParams = @{
                Path    = $env:TEMP + '\Excel.xlsx'
                Show    = $true
                Verbose = $true
            }
            Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
            Write-Output 489 668 299 777 860 151 119 497 234 788 |
                Export-Excel @ExcelParams -ConditionalText $(
                    New-ConditionalText -ConditionalType GreaterThan 525 -ConditionalTextColor DarkRed -BackgroundColor LightPink
                )

            Exports data that will have a Conditional Formatting rule in Excel
            that will show cells with a value is greater than 525, whith a
            background fill color of "LightPink" and the text in "DarkRed".
            Where condition is not met the color willbe the default, black
            text on a white background.

        .EXAMPLE
        >
        PS> $ExcelParams = @{
                Path    = $env:TEMP + '\Excel.xlsx'
                Show    = $true
                Verbose = $true
            }
            Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
            Get-Service | Select-Object -Property Name, Status, DisplayName, ServiceName |
                Export-Excel @ExcelParams -ConditionalText $(
                    New-ConditionalText Stop DarkRed LightPink
                    New-ConditionalText Running Blue Cyan
                )

            Exports all services to an Excel sheet, setting a Conditional formatting rule
            that will set the background fill color to "LightPink" and the text color
            to "DarkRed" when the value contains the word "Stop".
            If the value contains the word "Running" it will have a background fill
            color of "Cyan" and text colored 'Blue'. If neither condition is met, the
            color will be the default, black text on a white background.

        .EXAMPLE
        >
        PS> $ExcelParams = @{
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
            $Array | Update-FirstObjectProperties | Export-Excel @ExcelParams -WorksheetName Numbers

            Updates the first object of the array by adding property 'Member3' and 'Member4'.
            Afterwards. all objects are exported to an Excel file and all column headers are visible.

        .EXAMPLE
            Get-Process | Export-Excel .\test.xlsx -WorksheetName Processes -IncludePivotTable -Show -PivotRows Company -PivotData PM

        .EXAMPLE
            Get-Process | Export-Excel .\test.xlsx -WorksheetName Processes -ChartType PieExploded3D -IncludePivotChart -IncludePivotTable -Show -PivotRows Company -PivotData PM

        .EXAMPLE
            Get-Service | Export-Excel 'c:\temp\test.xlsx'  -Show -IncludePivotTable -PivotRows status -PivotData @{status='count'}

        .EXAMPLE
        >
        PS> $pt = [ordered]@{}
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
            Get-Process | Select-Object    -Property Name,Company,Handles,CPU,VM       | Export-Excel -Path .\test.xlsx -AutoSize -WorksheetName 'sheet2'
            Export-Excel -Path .\test.xlsx -PivotTableDefinition $pt -Show

            This example defines two PivotTables. Then it puts Service data on Sheet1
            with one call to Export-Excel and Process Data on sheet2 with a second
            call to Export-Excel. The third and final call adds the two PivotTables
            and opens the spreadsheet in Excel.
        .EXAMPLE
        >
        PS> Remove-Item  -Path .\test.xlsx
            $excel = Get-Service | Select-Object -Property Status,Name,DisplayName,StartType | Export-Excel -Path .\test.xlsx -PassThru
            $excel.Workbook.Worksheets["Sheet1"].Row(1).style.font.bold = $true
            $excel.Workbook.Worksheets["Sheet1"].Column(3 ).width = 29
            $excel.Workbook.Worksheets["Sheet1"].Column(3 ).Style.wraptext = $true
            $excel.Save()
            $excel.Dispose()
            Start-Process .\test.xlsx

            This example uses -PassThru. It puts service information into sheet1 of the
            workbook and saves the ExcelPackage object in $Excel. It then uses the package
            object to apply formatting, and then saves the workbook and disposes of the object
            before loading the document in Excel. Other commands in the module remove the need
            to work directly with the package object in this way.

        .EXAMPLE
        >
        PS> Remove-Item -Path .\test.xlsx -ErrorAction Ignore
            $excel = Get-Process | Select-Object -Property Name,Company,Handles,CPU,PM,NPM,WS | Export-Excel -Path .\test.xlsx -ClearSheet -WorksheetName "Processes" -PassThru
            $sheet = $excel.Workbook.Worksheets["Processes"]
            $sheet.Column(1) | Set-ExcelRange -Bold -AutoFit
            $sheet.Column(2) | Set-ExcelRange -Width 29 -WrapText
            $sheet.Column(3) | Set-ExcelRange -HorizontalAlignment Right -NFormat "#,###"
            Set-ExcelRange -Address $sheet.Cells["E1:H1048576"]  -HorizontalAlignment Right -NFormat "#,###"
            Set-ExcelRange -Address $sheet.Column(4)  -HorizontalAlignment Right -NFormat "#,##0.0" -Bold
            Set-ExcelRange -Address $sheet.Row(1) -Bold -HorizontalAlignment Center
            Add-ConditionalFormatting -WorkSheet $sheet -Range "D2:D1048576" -DataBarColor Red
            Add-ConditionalFormatting -WorkSheet $sheet -Range "G2:G1048576" -RuleType GreaterThan -ConditionValue "104857600" -ForeGroundColor Red
            foreach ($c in 5..9) {Set-ExcelRange -Address $sheet.Column($c)  -AutoFit }
            Export-Excel -ExcelPackage $excel -WorksheetName "Processes" -IncludePivotChart -ChartType ColumnClustered -NoLegend -PivotRows company  -PivotData @{'Name'='Count'}  -Show

            This a more sophisticated version of the previous example showing different
            ways of using Set-ExcelRange, and also adding conditional formatting.
            In the final command a PivotChart is added and the workbook is opened in Excel.
        .EXAMPLE
             0..360 | ForEach-Object {[pscustomobject][ordered]@{X=$_; Sinx="=Sin(Radians(x)) "} } | Export-Excel -now -LineChart -AutoNameRange

             Creates a line chart showing the value of Sine(x) for values of X between 0 and 360 degrees.

        .EXAMPLE
        >
        PS> Invoke-Sqlcmd -ServerInstance localhost\DEFAULT -Database AdventureWorks2014 -Query "select *  from sys.tables" -OutputAs DataRows |
            Export-Excel -Path .\SysTables_AdventureWorks2014.xlsx -WorksheetName Tables

            Runs a query against a SQL Server database and outputs the resulting rows DataRows using the -OutputAs parameter.
            The results are then piped to the Export-Excel function.
            NOTE: You need to install the SqlServer module from the PowerShell Gallery in oder to get the -OutputAs parameter for the Invoke-Sqlcmd cmdlet.

        .LINK
            https://github.com/dfinke/ImportExcel
    #>
    [CmdletBinding(DefaultParameterSetName = 'Now')]
    [OutputType([OfficeOpenXml.ExcelPackage])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "")]
    Param(

        [Parameter(Mandatory = $true, ParameterSetName = "Path", Position = 0)]
        [String]$Path,
        [Parameter(Mandatory = $true, ParameterSetName = "Package")]

        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        [Parameter(ValueFromPipeline = $true)]
        [Alias('TargetData')]
        $InputObject,
        [Switch]$Calculate,
        [Switch]$Show,
        [String]$WorksheetName = 'Sheet1',
        [String]$Password,
        [switch]$ClearSheet,
        [switch]$Append,
        [String]$Title,
        [OfficeOpenXml.Style.ExcelFillStyle]$TitleFillPattern = 'Solid',
        [Switch]$TitleBold,
        [Int]$TitleSize = 22,
        $TitleBackgroundColor,
        [Switch]$IncludePivotTable,
        [String]$PivotTableName,
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
        $MaxAutoSizeRows = 1000,
        [Switch]$NoClobber,
        [Switch]$FreezeTopRow,
        [Switch]$FreezeFirstColumn,
        [Switch]$FreezeTopRowFirstColumn,
        [Int[]]$FreezePane,


        [Switch]$AutoFilter,
        [Switch]$BoldTopRow,
        [Switch]$NoHeader,
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


        [String]$TableName,


        [OfficeOpenXml.Table.TableStyles]$TableStyle,
        [Switch]$Barchart,
        [Switch]$PieChart,
        [Switch]$LineChart ,
        [Switch]$ColumnChart ,
        [Object[]]$ExcelChartDefinition,
        [String[]]$HideSheet,
        [String[]]$UnHideSheet,
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
        [Switch]$NoAliasOrScriptPropeties,
        [Switch]$DisplayPropertySet,
        [String[]]$NoNumberConversion,
        [Object[]]$ConditionalFormat,
        [Object[]]$ConditionalText,
        [ScriptBlock]$CellStyleSB,
        #If there is already content in the workbook the sheet with the PivotTable will not be active UNLESS Activate is specified
        [switch]$Activate,
        [Parameter(ParameterSetName = 'Now')]
        [Switch]$Now,
        [Switch]$ReturnRange,
        #By default PivotTables have Totals for each Row (on the right) and for each column at the bottom. This allows just one or neither to be selected.
        [ValidateSet("Both","Columns","Rows","None")]
        [String]$PivotTotals = "Both",
        #Included for compatibility - equivalent to -PivotTotals "None"
        [Switch]$NoTotalsInPivot,
        [Switch]$ReZip
    )

    begin {
        $numberRegex = [Regex]'\d'
        $isDataTypeValueType = $false
        if ($NoClobber) {Write-Warning -Message "-NoClobber parameter is no longer used" }
        #Open the file, get the worksheet, and decide where in the sheet we are writing, and if there is a number format to apply.
        try   {
            $script:Header = $null
            if ($Append -and $ClearSheet) {throw "You can't use -Append AND -ClearSheet."}
            if ($PSBoundParameters.Keys.Count -eq 0 -Or $Now -or (-not $Path -and -not $ExcelPackage) ) {
                $Path = [System.IO.Path]::GetTempFileName() -replace '\.tmp', '.xlsx'
                $Show = $true
                $AutoSize = $true
                if (-not $TableName) {
                    $AutoFilter = $true
                }
            }
            if ($ExcelPackage) {
                $pkg = $ExcelPackage
                $Path = $pkg.File
            }
            Else { $pkg = Open-ExcelPackage -Path $Path -Create -KillExcel:$KillExcel -Password:$Password}
        }
        catch {throw "Could not open Excel Package $path"}
        try   {
            $params = @{WorksheetName=$WorksheetName}
            foreach ($p in @("ClearSheet", "MoveToStart", "MoveToEnd", "MoveBefore", "MoveAfter", "Activate")) {if ($PSBoundParameters[$p]) {$params[$p] = $PSBoundParameters[$p]}}
            $ws = $pkg | Add-WorkSheet @params
            if ($ws.Name -ne $WorksheetName) {
                Write-Warning -Message "The Worksheet name has been changed from $WorksheetName to $($ws.Name), this may cause errors later."
                $WorksheetName = $ws.Name
            }
        }
        catch {throw "Could not get worksheet $worksheetname"}
        try   {
            if ($Append -and $ws.Dimension) {
                #if there is a title or anything else above the header row, append needs to be combined wih a suitable startrow parameter
                $headerRange = $ws.Dimension.Address -replace "\d+$", $StartRow
                #using a slightly odd syntax otherwise header ends up as a 2D array
                $ws.Cells[$headerRange].Value | ForEach-Object -Begin {$Script:header = @()} -Process {$Script:header += $_ }
                #if we did not get AutoNameRange, but headers have ranges of the same name make autoNameRange True, otherwise make it false
                if (-not $AutoNameRange) {
                    $AutoNameRange  = $true ; foreach ($h in $header) {if ($ws.names.name -notcontains $h) {$AutoNameRange = $false} }
                }
                #if we did not get a Rangename but there is a Range which covers the active part of the sheet, set Rangename to that.
                if (-not $RangeName -and $ws.names.where({$_.name[0] -match '[a-z]'})) {
                    $theRange = $ws.names.where({
                         ($_.Name[0]   -match '[a-z]' )              -and
                         ($_.Start.Row    -eq $StartRow)             -and
                         ($_.Start.Column -eq $StartColumn)          -and
                         ($_.End.Row      -eq $ws.Dimension.End.Row) -and
                         ($_.End.Column   -eq $ws.Dimension.End.column) } , 'First', 1)
                    if ($theRange) {$rangename = $theRange.name}
                }

                #if we did not get a table name but there is a table which covers the active part of the sheet, set table name to that, and don't do anything with autofilter
                if (-not $TableName -and $ws.Tables.Where({$_.address.address -eq $ws.dimension.address})) {
                    $TableName  = $ws.Tables.Where({$_.address.address -eq $ws.dimension.address},'First', 1).Name
                    $AutoFilter = $false
                }
                #if we did not get $autofilter but a filter range is set and it covers the right area, set autofilter to true
                elseif (-not $AutoFilter -and $ws.Names['_xlnm._FilterDatabase']) {
                    if ( ($ws.Names['_xlnm._FilterDatabase'].Start.Row    -eq $StartRow)    -and
                         ($ws.Names['_xlnm._FilterDatabase'].Start.Column -eq $StartColumn) -and
                         ($ws.Names['_xlnm._FilterDatabase'].End.Row      -eq $ws.Dimension.End.Row) -and
                         ($ws.Names['_xlnm._FilterDatabase'].End.Column   -eq $ws.Dimension.End.Column) ) {$AutoFilter = $true}
                }

                $row = $ws.Dimension.End.Row
                Write-Debug -Message ("Appending: headers are " + ($script:Header -join ", ") + " Start row is $row")
                if ($Title) {Write-Warning -Message "-Title Parameter is ignored when appending."}
            }
            elseif ($Title) {
                #Can only add a title if not appending!
                $Row = $StartRow
                $ws.Cells[$Row, $StartColumn].Value = $Title
                $ws.Cells[$Row, $StartColumn].Style.Font.Size = $TitleSize

                if  ($PSBoundParameters.ContainsKey("TitleBold")) {
                    #Set title to Bold face font if -TitleBold was specified.
                    #Otherwise the default will be unbolded.
                    $ws.Cells[$Row, $StartColumn].Style.Font.Bold = [boolean]$TitleBold
                }
                if ($TitleBackgroundColor ) {
                    if ($TitleBackgroundColor -is [string])         {$TitleBackgroundColor = [System.Drawing.Color]::$TitleBackgroundColor }
                    $ws.Cells[$Row, $StartColumn].Style.Fill.PatternType = $TitleFillPattern
                    $ws.Cells[$Row, $StartColumn].Style.Fill.BackgroundColor.SetColor($TitleBackgroundColor)
                }
                $Row ++ ; $startRow ++
            }
            else {  $Row = $StartRow }
            $ColumnIndex = $StartColumn
            $Numberformat = Expand-NumberFormat -NumberFormat $Numberformat
            if ((-not $ws.Dimension) -and ($Numberformat -ne $ws.Cells.Style.Numberformat.Format)) {
                    $ws.Cells.Style.Numberformat.Format = $Numberformat
                    $setNumformat = $false
            }
            else {  $setNumformat = ($Numberformat -ne $ws.Cells.Style.Numberformat.Format) }
        }
        catch {throw "Failed preparing to export to worksheet '$WorksheetName' to '$Path': $_"}
        #region Special case -inputobject passed a dataTable object
        <# If inputObject was passed via the pipeline it won't be visible until the process block, we will only see it here if it was passed as a parameter
          if it was passed it is a data table don't do foreach on it (slow) put the whole table in and set dates on date columns,
          set things up for the end block, and skip the process block #>
        if ($InputObject -is  [System.Data.DataTable])  {
            $null = $ws.Cells[$row,$StartColumn].LoadFromDataTable($InputObject, (-not $noHeader) )
            foreach ($c in $InputObject.Columns.where({$_.datatype -eq [datetime]})) {
                Set-ExcelColumn -Worksheet $ws -Column ($c.Ordinal + $StartColumn) -NumberFormat 'Date-Time'
            }
            foreach ($c in $InputObject.Columns.where({$_.datatype -eq [timespan]})) {
                Set-ExcelColumn -Worksheet $ws -Column ($c.Ordinal + $StartColumn) -NumberFormat '[h]:mm:ss'
            }
            $ColumnIndex         += $InputObject.Columns.Count - 1
            if ($noHeader) {$row += $InputObject.Rows.Count -1 }
            else           {$row += $InputObject.Rows.Count    }
            $null = $PSBoundParameters.Remove('InputObject')
            $firstTimeThru = $false
        }
        #endregion
        else  {$firstTimeThru = $true}
    }

    process { if ($PSBoundParameters.ContainsKey("InputObject")) {
        try {
            if ($null -eq $InputObject) {$row += 1}
            foreach ($TargetData in $InputObject) {
                if ($firstTimeThru) {
                    $firstTimeThru = $false
                    $isDataTypeValueType = ($null -eq $TargetData) -or ($TargetData.GetType().name -match 'string|timespan|datetime|bool|byte|char|decimal|double|float|int|long|sbyte|short|uint|ulong|ushort|URI|ExcelHyperLink')
                    if ($isDataTypeValueType ) {
                        $script:Header = @(".")       # dummy value to make sure we go through the "for each name in $header"
                        if (-not $Append) {$row -= 1} # By default row will be 1, it is incremented before inserting values (so it ends pointing at final row.);  si first data row is 2 - move back up 1 if there is no header .
                    }
                    if ($null -ne $TargetData) {Write-Debug "DataTypeName is '$($TargetData.GetType().name)' isDataTypeValueType '$isDataTypeValueType'" }
                }
                #region Add headers - if we are appending, or we have been through here once already we will have the headers
                if (-not $script:Header) {
                    if ($DisplayPropertySet -and $TargetData.psStandardmembers.DefaultDisplayPropertySet.ReferencedPropertyNames) {
                        $script:Header = $TargetData.psStandardmembers.DefaultDisplayPropertySet.ReferencedPropertyNames.Where( {$_ -notin $ExcludeProperty})
                    }
                    else {
                        if ($NoAliasOrScriptPropeties) {$propType = "Property"} else {$propType = "*"}
                        $script:Header = $TargetData.PSObject.Properties.where( {$_.MemberType -like $propType}).Name
                    }
                    foreach ($exclusion in $ExcludeProperty) {$script:Header = $script:Header -notlike $exclusion}
                    if ($NoHeader) {
                        # Don't push the headers to the spreadsheet
                        $Row -= 1
                    }
                    else {
                        $ColumnIndex = $StartColumn
                        foreach ($Name in $script:Header) {
                            $ws.Cells[$Row, $ColumnIndex].Value = $Name
                            Write-Verbose "Cell '$Row`:$ColumnIndex' add header '$Name'"
                            $ColumnIndex += 1
                        }
                    }
                }
                #endregion
                #region Add non header values
                $Row += 1
                $ColumnIndex = $StartColumn
                <#
                 For each item in the header OR for the Data item if this is a simple Type or data table :
                   If it is a date insert with one of Excel's built in formats - recognized as "Date and time to be localized"
                   if it is a timespan insert with a built in format for elapsed hours, minutes and seconds
                   if its  any other numeric insert as is , setting format if need be.
                   Preserve URI, Insert a data table, convert non string objects to string.
                   For strings, check for fomula, URI or Number, before inserting as a string  (ignore nulls) #>
                foreach ($Name in $script:Header) {
                    if   ($isDataTypeValueType) {$v = $TargetData}
                    else {$v = $TargetData.$Name}
                    try   {
                        if     ($v -is    [DateTime]) {
                            $ws.Cells[$Row, $ColumnIndex].Value = $v
                            $ws.Cells[$Row, $ColumnIndex].Style.Numberformat.Format = 'm/d/yy h:mm' # This is not a custom format, but a preset recognized as date and localized.
                        }
                        elseif ($v -is    [TimeSpan]) {
                            $ws.Cells[$Row, $ColumnIndex].Value = $v
                            $ws.Cells[$Row, $ColumnIndex].Style.Numberformat.Format = '[h]:mm:ss'
                        }
                        elseif ($v -is    [System.ValueType]) {
                            $ws.Cells[$Row, $ColumnIndex].Value = $v
                            if ($setNumformat) {$ws.Cells[$Row, $ColumnIndex].Style.Numberformat.Format = $Numberformat }
                        }
                        elseif ($v -is    [uri] ) {
                            $ws.Cells[$Row, $ColumnIndex].HyperLink = $v
                            $ws.Cells[$Row, $ColumnIndex].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
                            $ws.Cells[$Row, $ColumnIndex].Style.Font.UnderLine = $true
                        }
                        elseif ($v -isnot [String] ) { #Other objects or null.
                            if ($null -ne $v) { $ws.Cells[$Row, $ColumnIndex].Value = $v.toString()}
                        }
                        elseif ($v[0] -eq '=') {
                            $ws.Cells[$Row, $ColumnIndex].Formula = ($v -replace '^=','')
                            if ($setNumformat) {$ws.Cells[$Row, $ColumnIndex].Style.Numberformat.Format = $Numberformat }
                        }
                        elseif ( [System.Uri]::IsWellFormedUriString($v , [System.UriKind]::Absolute) ) {
                            if ($v -match "^xl://internal/") {
                                  $referenceAddress = $v -replace "^xl://internal/" , ""
                                  $display          = $referenceAddress -replace "!A1$"   , ""
                                  $h = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $referenceAddress , $display
                                  $ws.Cells[$Row, $ColumnIndex].HyperLink = $h
                            }
                            else {$ws.Cells[$Row, $ColumnIndex].HyperLink = $v }   #$ws.Cells[$Row, $ColumnIndex].Value = $v.AbsoluteUri
                            $ws.Cells[$Row, $ColumnIndex].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
                            $ws.Cells[$Row, $ColumnIndex].Style.Font.UnderLine = $true
                        }
                        else {
                            $number = $null
                            if ( $numberRegex.IsMatch($v)     -and  # if it contains digit(s) - this syntax is quicker than -match for many items and cuts out slow checks for non numbers
                                 $NoNumberConversion -ne '*'  -and  # and NoNumberConversion isn't specified
                                 $NoNumberConversion -notcontains $Name -and
                                 [Double]::TryParse($v, [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::CurrentInfo, [Ref]$number)
                               ) {
                                 $ws.Cells[$Row, $ColumnIndex].Value = $number
                                 if ($setNumformat) {$ws.Cells[$Row, $ColumnIndex].Style.Numberformat.Format = $Numberformat }
                            }
                            else {
                                $ws.Cells[$Row, $ColumnIndex].Value  = $v
                            }

                        }
                    }
                    catch {Write-Warning -Message "Could not insert the '$Name' property at Row $Row, Column $ColumnIndex"}
                    $ColumnIndex += 1
                }
                $ColumnIndex -= 1 # column index will be the last column whether isDataTypeValueType was true or false
                #endregion
            }
        }
        catch {throw "Failed exporting data to worksheet '$WorksheetName' to '$Path': $_" }

    }}

    end {
        if ($firstTimeThru -and $ws.Dimension) {
              $LastRow        = $ws.Dimension.End.Row
              $LastCol        = $ws.Dimension.End.Column
              $endAddress     = $ws.Dimension.End.Address
        }
        else {
              $LastRow        = $Row
              $LastCol        = $ColumnIndex
              $endAddress     = [OfficeOpenXml.ExcelAddress]::GetAddress($LastRow , $LastCol)
        }
        $startAddress         = [OfficeOpenXml.ExcelAddress]::GetAddress($StartRow, $StartColumn)
        $dataRange            = "{0}:{1}" -f $startAddress, $endAddress

        Write-Debug "Data Range '$dataRange'"
        if ($AutoNameRange) {
            try {
                if (-not $script:header) {
                    # if there aren't any headers, use the the first row of data to name the ranges: this is the last point that headers will be used.
                    $headerRange = $ws.Dimension.Address -replace "\d+$", $StartRow
                    #using a slightly odd syntax otherwise header ends up as a 2D array
                    $ws.Cells[$headerRange].Value | ForEach-Object -Begin {$Script:header = @()} -Process {$Script:header += $_ }
                    if   ($PSBoundParameters.ContainsKey($TargetData)) {  #if Export was called with data that writes no header start the range at $startRow ($startRow is data)
                           $targetRow = $StartRow
                    }
                    else { $targetRow = $StartRow + 1 }                   #if Export was called without data to add names (assume $startRow is a header) or...
                }                                                         #          ... called with data that writes a header, then start the range at $startRow + 1
                else {     $targetRow = $StartRow + 1 }

                #Dimension.start.row always seems to be one so we work out the target row
                #, but start.column is the first populated one and .Columns is the count of populated ones.
                # if we have 5 columns from 3 to 8, headers are numbered 0..4, so that is in the for loop and used for getting the name...
                # but we have to add the start column on when referencing positions
                foreach ($c in 0..($LastCol - $StartColumn)) {
                    $targetRangeName = @($script:Header)[$c]  #Let Add-ExcelName fix (and warn about) bad names
                    Add-ExcelName  -RangeName $targetRangeName -Range $ws.Cells[$targetRow, ($StartColumn + $c ), $LastRow, ($StartColumn + $c )]
                    try {#this test can throw with some names, surpress any error
                        if ([OfficeOpenXml.FormulaParsing.ExcelUtilities.ExcelAddressUtil]::IsValidAddress(($targetRangeName -replace '\W' , '_' ))) {
                            Write-Warning -Message "AutoNameRange: Property name '$targetRangeName' is also a valid Excel address and may cause issues. Consider renaming the property."
                        }
                    }
                    Catch {
                        Write-Warning -Message "AutoNameRange: Testing '$targetRangeName' caused an error. This should be harmless, but a change of property name may be needed.."
                    }
                }
            }
            catch {Write-Warning -Message "Failed adding named ranges to worksheet '$WorksheetName': $_"  }
        }
        #Empty string is not allowed as a name for ranges or tables.
        if ($RangeName) { Add-ExcelName  -Range $ws.Cells[$dataRange] -RangeName $RangeName}

        #Allow table to be inserted by specifying Name, or Style or both; only process autoFilter if there is no table (they clash).
        if     ($TableName) {
            if ($PSBoundParameters.ContainsKey('TableStyle')) {
                  Add-ExcelTable -Range $ws.Cells[$dataRange] -TableName $TableName -TableStyle $TableStyle
            }
            else {Add-ExcelTable -Range $ws.Cells[$dataRange] -TableName $TableName}
        }
        elseif ($PSBoundParameters.ContainsKey('TableStyle')) {
                  Add-ExcelTable -Range $ws.Cells[$dataRange] -TableName "" -TableStyle $TableStyle
        }
        elseif ($AutoFilter) {
            try {
                $ws.Cells[$dataRange].AutoFilter = $true
                Write-Verbose -Message "Enabled autofilter. "
            }
            catch {Write-Warning -Message "Failed adding autofilter to worksheet '$WorksheetName': $_"}
        }

        if ($PivotTableDefinition) {
            foreach ($item in $PivotTableDefinition.GetEnumerator()) {
                $params = $item.value
                if ($Activate) {$params.Activate = $true   }
                if ($params.keys -notcontains 'SourceRange' -and
                   ($params.Keys -notcontains 'SourceWorkSheet'   -or  $params.SourceWorkSheet -eq $WorksheetName)) {$params.SourceRange = $dataRange}
                if ($params.Keys -notcontains 'SourceWorkSheet')      {$params.SourceWorkSheet = $ws }
                if ($params.Keys -notcontains 'NoTotalsInPivot'   -and $NoTotalsInPivot  ) {$params.PivotTotals       = 'None'}
                if ($params.Keys -notcontains 'PivotTotals'       -and $PivotTotals      ) {$params.PivotTotals       = $PivotTotals}
                if ($params.Keys -notcontains 'PivotDataToColumn' -and $PivotDataToColumn) {$params.PivotDataToColumn = $true}

                Add-PivotTable -ExcelPackage $pkg -PivotTableName $item.key @Params
            }
        }
        if ($IncludePivotTable -or $IncludePivotChart) {
            $params = @{
                'SourceRange' = $dataRange
            }
            if ($PivotTableName -and ($pkg.workbook.worksheets.tables.name -contains $PivotTableName)) {
                Write-Warning -Message "The selected PivotTable name '$PivotTableName' is already used as a table name. Adding a suffix of 'Pivot'."
                $PivotTableName += 'Pivot'
            }

            if   ($PivotTableName)  {$params.PivotTableName    = $PivotTableName}
            else                    {$params.PivotTableName    = $WorksheetName + 'PivotTable'}
            if          ($Activate) {$params.Activate          = $true   }
            if       ($PivotFilter) {$params.PivotFilter       = $PivotFilter}
            if         ($PivotRows) {$params.PivotRows         = $PivotRows}
            if      ($PivotColumns) {$Params.PivotColumns      = $PivotColumns}
            if         ($PivotData) {$Params.PivotData         = $PivotData}
            if   ($NoTotalsInPivot) {$params.PivotTotals       = "None"    }
            Elseif   ($PivotTotals) {$params.PivotTotals       = $PivotTotals}
            if ($PivotDataToColumn) {$params.PivotDataToColumn = $true}
            if ($IncludePivotChart) {
                                     $params.IncludePivotChart = $true
                                     $Params.ChartType         = $ChartType
                if ($ShowCategory)  {$params.ShowCategory      = $true}
                if ($ShowPercent)   {$params.ShowPercent       = $true}
                if ($NoLegend)      {$params.NoLegend          = $true}
            }
            Add-PivotTable -ExcelPackage $pkg -SourceWorkSheet $ws   @params
        }

        try {
            #Allow single switch or two seperate ones.
            if ($FreezeTopRowFirstColumn -or ($FreezeTopRow -and $FreezeFirstColumn)) {
                $ws.View.FreezePanes(2, 2)
                Write-Verbose -Message "Froze top row and first column"
            }
            elseif ($FreezeTopRow) {
                $ws.View.FreezePanes(2, 1)
                Write-Verbose -Message "Froze top row"
            }
            elseif ($FreezeFirstColumn) {
                $ws.View.FreezePanes(1, 2)
                Write-Verbose -Message "Froze first column"
            }
            #Must be 1..maxrows or and array of 1..maxRows,1..MaxCols
            if ($FreezePane) {
                $freezeRow, $freezeColumn = $FreezePane
                if (-not $freezeColumn -or $freezeColumn -eq 0) {
                    $freezeColumn = 1
                }

                if ($freezeRow -ge 1) {
                    $ws.View.FreezePanes($freezeRow, $freezeColumn)
                    Write-Verbose -Message "Froze panes at row $freezeRow and column $FreezeColumn"
                }
            }
        }
        catch {Write-Warning -Message "Failed adding Freezing the panes in worksheet '$WorksheetName': $_"}

        if  ($PSBoundParameters.ContainsKey("BoldTopRow")) { #it sets bold as far as there are populated cells: for whole row could do $ws.row($x).style.font.bold = $true
            try {
                if ($Title) {
                    $range = $ws.Dimension.Address -replace '\d+', ($StartRow + 1)
                }
                else {
                    $range = $ws.Dimension.Address -replace '\d+', $StartRow
                }
                $ws.Cells[$range].Style.Font.Bold = [boolean]$BoldTopRow
                Write-Verbose -Message "Set $range font style to bold."
            }
            catch {Write-Warning -Message "Failed setting the top row to bold in worksheet '$WorksheetName': $_"}
        }
        if ($AutoSize) {
            try {
                #Don't fit the all the columns in the sheet; if we are adding cells beside things with hidden columns, that unhides them
                if ($MaxAutoSizeRows -and $MaxAutoSizeRows -lt $LastRow ) {
                    $AutosizeRange = [OfficeOpenXml.ExcelAddress]::GetAddress($startRow,$StartColumn,   $MaxAutoSizeRows , $LastCol)
                    $ws.Cells[$AutosizeRange].AutoFitColumns()
                }
                else {$ws.Cells[$dataRange].AutoFitColumns()  }
                Write-Verbose -Message "Auto-sized columns"
            }
            catch {  Write-Warning -Message "Failed autosizing columns of worksheet '$WorksheetName': $_"}
        }

        foreach ($Sheet in $HideSheet) {
            try {
                $pkg.Workbook.WorkSheets.Where({$_.Name -like $sheet}) | ForEach-Object {
                    $_.Hidden = 'Hidden'
                    Write-verbose -Message "Sheet '$($_.Name)' Hidden."
                }
            }
            catch {Write-Warning -Message  "Failed hiding worksheet '$sheet': $_"}
        }
        foreach ($Sheet in $UnHideSheet) {
            try {
                $pkg.Workbook.WorkSheets.Where({$_.Name -like $sheet}) | ForEach-Object {
                    $_.Hidden = 'Visible'
                    Write-verbose -Message "Sheet '$($_.Name)' shown"
                }
            }
            catch {Write-Warning -Message  "Failed showing worksheet '$sheet': $_"}
        }
        if (-not $pkg.Workbook.Worksheets.Where({$_.Hidden -eq 'visible'})) {
            Write-Verbose -Message "No Sheets were left visible, making $WorksheetName visible"
            $ws.Hidden = 'Visible'
        }

        foreach ($chartDef in $ExcelChartDefinition) {
            if ($chartDef -is [System.Management.Automation.PSCustomObject]) {
                $params = @{}
                $chartDef.PSObject.Properties | ForEach-Object {if ( $null -ne $_.value) {$params[$_.name] = $_.value}}
                Add-ExcelChart -Worksheet $ws @params
            }
            elseif ($chartDef -is [hashtable] -or  $chartDef -is[System.Collections.Specialized.OrderedDictionary]) {
                Add-ExcelChart -Worksheet $ws @chartDef
            }
        }

        if ($Calculate) {
            try   { [OfficeOpenXml.CalculationExtension]::Calculate($ws) }
            catch { Write-Warning "One or more errors occured while calculating, save will continue, but there may be errors in the workbook. $_"}
        }

        if ($Barchart -or $PieChart -or $LineChart -or $ColumnChart) {
            if ($NoHeader) {$FirstDataRow = $startRow}
            else           {$FirstDataRow = $startRow + 1 }
            $range = [OfficeOpenXml.ExcelAddress]::GetAddress($FirstDataRow, $startColumn, $FirstDataRow, $lastCol )
            $xCol  = $ws.cells[$range] | Where-Object {$_.value -is [string]    } | ForEach-Object {$_.start.column} | Sort-Object | Select-Object -first 1
            if (-not $xcol) {
                $xcol  = $StartColumn
                $range = [OfficeOpenXml.ExcelAddress]::GetAddress($FirstDataRow, ($startColumn +1), $FirstDataRow, $lastCol )
            }
            $yCol  = $ws.cells[$range] | Where-Object {$_.value -is [valueType] -or $_.Formula } | ForEach-Object {$_.start.column} | Sort-Object | Select-Object -first 1
            if (-not ($xCol -and $ycol)) { Write-Warning -Message "Can't identify a string column and a number column to use as chart labels and data. "}
            else {
                $params = @{
                XRange = [OfficeOpenXml.ExcelAddress]::GetAddress($FirstDataRow, $xcol , $lastrow, $xcol)
                YRange = [OfficeOpenXml.ExcelAddress]::GetAddress($FirstDataRow, $ycol , $lastrow, $ycol)
                Title  =  ''
                Column = ($lastCol +1)
                Width  = 800
                }
                if   ($ShowPercent) {$params["ShowPercent"]  = $true}
                if  ($ShowCategory) {$params["ShowCategory"] = $true}
                if      ($NoLegend) {$params["NoLegend"]     = $true}
                if (-not $NoHeader) {$params["SeriesHeader"] = $ws.Cells[$startRow, $YCol].Value}
                if   ($ColumnChart) {$Params["chartType"]    = "ColumnStacked" }
                elseif  ($Barchart) {$Params["chartType"]    = "BarStacked"    }
                elseif  ($PieChart) {$Params["chartType"]    = "PieExploded3D" }
                elseif ($LineChart) {$Params["chartType"]    = "Line"          }

                Add-ExcelChart -Worksheet $ws @params
            }
        }

        # It now doesn't matter if the conditional formating rules are passed in $conditionalText or $conditional format.
        # Just one with an alias for compatiblity it will break things for people who are using both at once
        foreach ($c in  (@() + $ConditionalText  +  $ConditionalFormat) ) {
            try {
                #we can take an object with a .ConditionalType property made by New-ConditionalText or with a .Formatter Property made by New-ConditionalFormattingIconSet or a hash table
                if ($c.ConditionalType) {
                    $cfParams = @{RuleType = $c.ConditionalType;    ConditionValue = $c.Text ;
                           BackgroundColor = $c.BackgroundColor; BackgroundPattern = $c.PatternType  ;
                           ForeGroundColor = $c.ConditionalTextColor}
                    if ($c.Range) {$cfParams.Range = $c.Range}
                    else          {$cfParams.Range = $ws.Dimension.Address }
                    Add-ConditionalFormatting -WorkSheet $ws @cfParams
                    Write-Verbose -Message "Added conditional formatting to range $($c.range)"
                }
                elseif ($c.formatter)  {
                    switch ($c.formatter) {
                        "ThreeIconSet" {Add-ConditionalFormatting -WorkSheet $ws -ThreeIconsSet $c.IconType -range $c.range -reverse:$c.reverse  }
                        "FourIconSet"  {Add-ConditionalFormatting -WorkSheet $ws  -FourIconsSet $c.IconType -range $c.range -reverse:$c.reverse  }
                        "FiveIconSet"  {Add-ConditionalFormatting -WorkSheet $ws  -FiveIconsSet $c.IconType -range $c.range -reverse:$c.reverse  }
                    }
                    Write-Verbose -Message "Added conditional formatting to range $($c.range)"
                }
                elseif ($c -is [hashtable] -or  $c -is[System.Collections.Specialized.OrderedDictionary]) {
                    if (-not $c.Range -or $c.Address) {$c.Address = $ws.Dimension.Address }
                    Add-ConditionalFormatting -WorkSheet $ws @c
                }
            }
            catch {throw "Error applying conditional formatting to worksheet $_"}
        }

        if ($CellStyleSB) {
            try {
                $TotalRows = $ws.Dimension.Rows
                $LastColumn = $ws.Dimension.Address -replace "^.*:(\w*)\d+$" , '$1'
                & $CellStyleSB $ws $TotalRows $LastColumn
            }
            catch {Write-Warning -Message "Failed processing CellStyleSB in worksheet '$WorksheetName': $_"}
        }

        #Can only add password, may want to support -password $Null removing password.
        if ($Password) {
            try {
                $ws.Protection.SetPassword($Password)
                Write-Verbose -Message 'Set password on workbook'
            }

            catch {throw "Failed setting password for worksheet '$WorksheetName': $_"}
        }

        if ($PassThru) {       $pkg   }
        else {
            if ($ReturnRange) {$dataRange }

            if ($Password) { $pkg.Save($Password) }
            else           { $pkg.Save() }
            Write-Verbose -Message "Saved workbook $($pkg.File)"
            if ($ReZip) {
                Write-Verbose -Message "Re-Zipping $($pkg.file) using .NET ZIP library"
                try {
                    Add-Type -AssemblyName 'System.IO.Compression.Filesystem' -ErrorAction stop
                }
                catch {
                    Write-Error "The -ReZip parameter requires .NET Framework 4.5 or later to be installed. Recommend to install Powershell v4+"
                    continue
                }
                try {
                    $TempZipPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
                    $null = [io.compression.zipfile]::ExtractToDirectory($pkg.File, $TempZipPath)
                    Remove-Item $pkg.File -Force
                    $null = [io.compression.zipfile]::CreateFromDirectory($TempZipPath, $pkg.File)
                }
                catch {throw "Error resizipping $path : $_"}
            }

            $pkg.Dispose()

            if ($Show) { Invoke-Item $Path }
        }

    }
}

function Add-WorkSheet  {
    <#
      .Synopsis
        Adds a worksheet to an existing workbook.
      .Description
        If the named worksheet already exists, the -Clearsheet parameter decides whether it should be deleted and a new one returned,
        or if not specified the existing sheet will be returned. By default the sheet is created at the end of the work book, the
        -MoveXXXX switches allow the sheet to be [re]positioned at the start or before or after another sheet. A new sheet will only be
        made the default sheet when excel opens if -Activate is specified.
      .Example
        $WorksheetActors = $ExcelPackage | Add-WorkSheet -WorkSheetname Actors

        $ExcelPackage holds an Excel package object (returned by Open-ExcelPackage, or Export-Excel -passthru).
        This command will add a sheet named 'Actors', or return the sheet if it exists, and the result is stored in $WorkSheetActors.
      .Example
        $WorksheetActors = Add-WorkSheet -ExcelPackage $ExcelPackage -WorkSheetname "Actors" -ClearSheet -MoveToStart

        This time the Excel package object is passed as a parameter instead of piped. If the 'Actors' sheet already exists it is deleted
        and  re-created. The new sheet will be created last in the workbook, and -MoveToStart Moves it to the start.
      .Example
        $null = Add-WorkSheet -ExcelWorkbook $wb -WorkSheetname $DestinationName -CopySource  $sourceWs -Activate
        This time a workbook is used instead of a package, and a worksheet is copied - $SourceWs is a worksheet object, which can come
        from the same workbook or a different one. Here the new copy of the data is made the active sheet when the workbook is opened.
    #>
    [cmdletBinding()]
    [OutputType([OfficeOpenXml.ExcelWorksheet])]
    param(
        #An object representing an Excel Package.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "Package", Position = 0)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        #An Excel Workbook to which the Worksheet will be added - a Package contains one Workbook, so you can use whichever fits at the time.
        [Parameter(Mandatory = $true, ParameterSetName = "WorkBook")]
        [OfficeOpenXml.ExcelWorkbook]$ExcelWorkbook,
        #The name of the worksheet, 'Sheet1' by default.
        [string]$WorksheetName ,
        #If the worksheet already exists, by default it will returned, unless -ClearSheet is specified in which case it will be deleted and re-created.
        [switch]$ClearSheet,
        #If specified, the worksheet will be moved to the start of the workbook.
        #MoveToStart takes precedence over MoveToEnd, Movebefore and MoveAfter if more than one is specified.
        [Switch]$MoveToStart,
        #If specified, the worksheet will be moved to the end of the workbook.
        #(This is the default position for newly created sheets, but this can be used to move existing sheets.)
        [Switch]$MoveToEnd,
        #If specified, the worksheet will be moved before the nominated one (which can be an index starting from 1, or a name).
        #MoveBefore takes precedence over MoveAfter if both are specified.
        $MoveBefore ,
        # If specified, the worksheet will be moved after the nominated one (which can be an index starting from 1, or a name or *).
        # If * is used, the worksheet names will be examined starting with the first one, and the sheet placed after the last sheet which comes before it alphabetically.
        $MoveAfter ,
        #If there is already content in the workbook the new sheet will not be active UNLESS Activate is specified.
        [switch]$Activate,
        #If worksheet is provided as a copy source the new worksheet will be a copy of it. The source can be in the same workbook, or in a different file.
        [OfficeOpenXml.ExcelWorksheet]$CopySource,
        #Ignored but retained for backwards compatibility.
        [Switch] $NoClobber
    )
    #if we were given a workbook use it, if we were given a package, use its workbook
    if      ($ExcelPackage -and -not $ExcelWorkbook) {$ExcelWorkbook = $ExcelPackage.Workbook}

    # If WorksheetName was given, try to use that worksheet. If it wasn't, and we are copying an existing sheet, try to use the sheet name
    # If we are not copying a sheet, and have no name, use the name "SheetX" where X is the number of the new sheet
    if      (-not $WorksheetName -and $CopySource -and -not $ExcelWorkbook[$CopySource.Name]) {$WorksheetName = $CopySource.Name}
    elseif  (-not $WorksheetName) {$WorksheetName = "Sheet" + (1 + $ExcelWorkbook.Worksheets.Count)}
    else    {$ws = $ExcelWorkbook.Worksheets[$WorksheetName]}

    #If -clearsheet was specified and the named sheet exists, delete it
    if      ($ws -and $ClearSheet) { $ExcelWorkbook.Worksheets.Delete($WorksheetName) ; $ws = $null }

    #Copy or create new sheet as needed
    if (-not $ws -and $CopySource) {
          Write-Verbose -Message "Copying into worksheet '$WorksheetName'."
          $ws = $ExcelWorkbook.Worksheets.Add($WorksheetName, $CopySource)
    }
    elseif (-not $ws) {
          $ws = $ExcelWorkbook.Worksheets.Add($WorksheetName)
          Write-Verbose -Message "Adding worksheet '$WorksheetName'."
    }
    else {Write-Verbose -Message "Worksheet '$WorksheetName' already existed."}
    #region Move sheet if needed
    if     ($MoveToStart) {$ExcelWorkbook.Worksheets.MoveToStart($WorksheetName) }
    elseif ($MoveToEnd  ) {$ExcelWorkbook.Worksheets.MoveToEnd($WorksheetName)   }
    elseif ($MoveBefore ) {
        if ($ExcelWorkbook.Worksheets[$MoveBefore]) {
            if ($MoveBefore -is [int]) {
                $ExcelWorkbook.Worksheets.MoveBefore($ws.Index, $MoveBefore)
            }
            else {$ExcelWorkbook.Worksheets.MoveBefore($WorksheetName, $MoveBefore)}
        }
        else {Write-Warning "Can't find worksheet '$MoveBefore'; worsheet '$WorksheetName' will not be moved."}
    }
    elseif ($MoveAfter  ) {
        if ($MoveAfter -eq "*") {
            if ($WorksheetName -lt $ExcelWorkbook.Worksheets[1].Name) {$ExcelWorkbook.Worksheets.MoveToStart($WorksheetName)}
            else {
                $i = 1
                While ($i -lt $ExcelWorkbook.Worksheets.Count -and ($ExcelWorkbook.Worksheets[$i + 1].Name -le $WorksheetName) ) { $i++}
                $ExcelWorkbook.Worksheets.MoveAfter($ws.Index, $i)
            }
        }
        elseif ($ExcelWorkbook.Worksheets[$MoveAfter]) {
            if ($MoveAfter -is [int]) {
                $ExcelWorkbook.Worksheets.MoveAfter($ws.Index, $MoveAfter)
            }
            else {
                $ExcelWorkbook.Worksheets.MoveAfter($WorksheetName, $MoveAfter)
            }
        }
        else {Write-Warning "Can't find worksheet '$MoveAfter'; worsheet '$WorksheetName' will not be moved."}
    }
    #endregion
    if ($Activate) {Select-Worksheet -ExcelWorksheet $ws  }
    if ($ExcelPackage -and -not (Get-Member -InputObject $ExcelPackage -Name $ws.Name)) {
        $sb = [scriptblock]::Create(('$this.workbook.Worksheets["{0}"]' -f $ws.name))
        Add-Member -InputObject $ExcelPackage -MemberType ScriptProperty -Name $ws.name -Value $sb
    }
    return $ws
}

function Select-Worksheet {
   <#
      .SYNOPSIS
        Sets the selected tab in an Excel workbook to be the chosen sheet and unselects all the others.
      .DESCRIPTION
        Sometimes when a sheet is added we want it to be the active sheet, sometimes we want the active sheet to be left as it was.
        Select-Worksheet exists to change which sheet is the selected tab when Excel opens the file.
      .EXAMPLE
        Select-Worksheet -ExcelWorkbook $ExcelWorkbook -WorksheetName "NewSheet"
        $ExcelWorkbook holds a workbook object containing a sheet named "NewSheet";
        This sheet will become the [only] active sheet in the workbook
      .EXAMPLE
        Select-Worksheet -ExcelPackage $Pkg -WorksheetName "NewSheet2"
        $pkg holds an Excel Package, whose workbook contains a sheet named "NewSheet2"
        This sheet will become the [only] active sheet in the workbook.
      .EXAMPLE
        Select-Worksheet -ExcelWorksheet $ws
        $ws holds an Excel worksheet which will become the [only] active sheet
        in its workbook.
    #>
    param (
        #An object representing an ExcelPackage.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'Package', Position = 0)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        #An Excel workbook to which the Worksheet will be added - a package contains one Workbook so you can use workbook or package as it suits.
        [Parameter(Mandatory = $true, ParameterSetName = 'WorkBook')]
        [OfficeOpenXml.ExcelWorkbook]$ExcelWorkbook,
        [Parameter(ParameterSetName='Package')]
        [Parameter(ParameterSetName='Workbook')]
        #The name of the worksheet "Sheet1" by default.
        [string]$WorksheetName,
        #An object representing an Excel worksheet.
        [Parameter(ParameterSetName='Sheet',Mandatory=$true)]
        [OfficeOpenXml.ExcelWorksheet]$ExcelWorksheet
    )
    #if we were given a package, use its workbook
    if      ($ExcelPackage   -and -not $ExcelWorkbook) {$ExcelWorkbook  = $ExcelPackage.Workbook}
    #if we now have workbook, get the worksheet; if we were given a sheet get the workbook
    if      ($ExcelWorkbook  -and $WorksheetName)      {$ExcelWorksheet = $ExcelWorkbook.Worksheets[$WorksheetName]}
    elseif  ($ExcelWorksheet -and -not $ExcelWorkbook) {$ExcelWorkbook  = $ExcelWorksheet.Workbook ; }
    #if we didn't get to a worksheet give up. If we did set all works sheets to not selected and then the one we want to selected.
    if (-not $ExcelWorksheet) {Write-Warning -Message "The worksheet $WorksheetName was not found." ; return }
    else {
        foreach ($w in $ExcelWorkbook.Worksheets) {$w.View.TabSelected = $false}
        $ExcelWorksheet.View.TabSelected = $true
    }
}

function Add-ExcelName {
    <#
      .SYNOPSIS
        Adds a named-range to an existing Excel worksheet.
      .DESCRIPTION
        It is often helpful to be able to refer to sets of cells with a name rather than using their co-ordinates; Add-ExcelName sets up these names.
      .EXAMPLE
          Add-ExcelName -Range $ws.Cells[$dataRange] -RangeName $rangeName
          $WS is a worksheet, and $dataRange is a string describing a range of cells - e.g. "A1:Z10"
          which will become a named range, using the name in $rangeName.
    #>
    [CmdletBinding()]
    param(
        #The range of cells to assign as a name.
        [Parameter(Mandatory=$true)]
        [OfficeOpenXml.ExcelRange]$Range,
        #The name to assign to the range. If the name exists it will be updated to the new range. If no name is specified, the first cell in the range will be used as the name.
        [String]$RangeName
    )
    try {
        $ws = $Range.Worksheet
        if (-not $RangeName) {
            $RangeName = $ws.Cells[$Range.Start.Address].Value
            $Range  = ($Range.Worksheet.cells[($range.start.row +1), $range.start.Column ,  $range.end.row, $range.end.column])
        }
        if ($RangeName -match '\W') {
            Write-Warning -Message "Range name '$RangeName' contains illegal characters, they will be replaced with '_'."
            $RangeName = $RangeName -replace '\W','_'
        }
        if ($ws.names[$RangeName]) {
            Write-verbose -Message "Updating Named range '$RangeName' to $($Range.FullAddressAbsolute)."
            $ws.Names[$RangeName].Address = $Range.FullAddressAbsolute
        }
        else  {
            Write-verbose -Message "Creating Named range '$RangeName' as $($Range.FullAddressAbsolute)."
            $null = $ws.Names.Add($RangeName, $Range)
        }
    }
    catch {Write-Warning -Message "Failed adding named range '$RangeName' to worksheet '$($ws.Name)': $_"  }
}

function Add-ExcelTable {
    <#
      .SYNOPSIS
        Adds Tables to Excel workbooks.
      .DESCRIPTION
        Unlike named ranges, where the name only needs to be unique within a sheet, Table names must be unique in the workbook
        Tables carry formatting by default have a filter. The filter, header, Totals, first and last column highlights
      .EXAMPLE
        Add-ExcelTable -Range $ws.Cells[$dataRange] -TableName $TableName

        $WS is a worksheet, and $dataRange is a string describing a range of cells - e.g. "A1:Z10"
        this range which will become a table, named $TableName
      .EXAMPLE
        Add-ExcelTable -Range $ws.cells[$($ws.Dimension.address)] -TableStyle Light1 -TableName Musictable -ShowFilter:$false -ShowTotal -ShowFirstColumn
        Again $ws is a worksheet, range here is the whole of the active part of the worksheet. The table style and name are set,
        the filter is turned off, and a "Totals" row added, and first column is set in bold.
    #>
    [CmdletBinding()]
    [OutputType([OfficeOpenXml.Table.ExcelTable])]
    param (
        #The range of cells to assign to a table.
        [Parameter(Mandatory=$true)]
        [OfficeOpenXml.ExcelRange]$Range,
        #The name for the Table - this should be unqiue in the Workbook - auto generated names will be used if this is left empty.
        [String]$TableName = "",
        #The Style for the table, by default "Medium6" is used
        [OfficeOpenXml.Table.TableStyles]$TableStyle = 'Medium6',
        #By default the header row is shown - it can be turned off with -ShowHeader:$false.
        [Switch]$ShowHeader ,
        #By default the filter is enabled - it can be turned off with -ShowFilter:$false.
        [Switch]$ShowFilter,
        #Show total adds a totals row. This does not automatically sum the columns but provides a drop-down in each to select sum, average etc
        [Switch]$ShowTotal,
        #A HashTable in the form ColumnName = "Average"|"Count"|"CountNums"|"Max"|"Min"|"None"|"StdDev"|"Sum"|"Var" - if specified, -ShowTotal is not needed.
        [hashtable]$TotalSettings,
        #Highlights the first column in bold.
        [Switch]$ShowFirstColumn,
        #Highlights the last column in bold.
        [Switch]$ShowLastColumn,
        #By default the table formats show striped rows, the can be turned off with -ShowRowStripes:$false
        [Switch]$ShowRowStripes,
        #Turns on column stripes.
        [Switch]$ShowColumnStripes,
        #If -PassThru is specified, the table object will be returned to allow additional changes to be made.
        [Switch]$PassThru
    )
    try {
        if ($TableName -eq "" -or $null -eq $TableName) {
            $tbl = $Range.Worksheet.Tables.Add($Range, "")
        }
        else {
            if ([OfficeOpenXml.FormulaParsing.ExcelUtilities.ExcelAddressUtil]::IsValidAddress($TableName)) {
                Write-Warning -Message "$TableName reads as an Excel address, and so is not allowed as a table name."
                return
            }
            if ($TableName -notMatch '^[A-Z]') {
                Write-Warning -Message "$TableName is not allowed as a table name because it does not begin with a letter."
                return
            }
            if ($TableName -match "\W") {
                Write-Warning -Message "At least one character in $TableName is illegal in a table name and will be replaced with '_' . "
                $TableName = $TableName -replace '\W', '_'
            }
            $ws = $Range.Worksheet
            #if the table exists in this worksheet, update it.
            if ($ws.Tables[$TableName]) {
                $tbl =$ws.Tables[$TableName]
                $tbl.TableXml.table.ref = $Range.Address
                Write-Verbose -Message "Re-defined table '$TableName', now at $($Range.Address)."
            }
            elseif ($ws.Workbook.Worksheets.Tables.Name -contains $TableName) {
                Write-Warning -Message "The Table name '$TableName' is already used on a different worksheet."
                return
            }
            else {
                $tbl = $ws.Tables.Add($Range, $TableName)
                Write-Verbose -Message "Defined table '$($tbl.Name)' at $($Range.Address)"
            }
        }
        #it seems that show total changes some of the others, so the sequence matters.
        if     ($PSBoundParameters.ContainsKey('ShowHeader'))        {$tbl.ShowHeader        = [bool]$ShowHeader}
        if     ($PSBoundParameters.ContainsKey('TotalSettings'))     {
            $tbl.ShowTotal = $true
            foreach ($k in $TotalSettings.keys) {
                if (-not $tbl.Columns[$k]) {Write-Warning -Message "Table does not have a Column '$k'."}
                elseif ($TotalSettings[$k] -notin @("Average", "Count", "CountNums", "Max", "Min", "None", "StdDev", "Sum", "Var") ) {
                    Write-Warning -Message "'$($TotalSettings[$k])' is not a valid total function."
                }
                else {$tbl.Columns[$k].TotalsRowFunction = $TotalSettings[$k]}
            }
        }
        elseif ($PSBoundParameters.ContainsKey('ShowTotal'))         {$tbl.ShowTotal         = [bool]$ShowTotal}
        if     ($PSBoundParameters.ContainsKey('ShowFilter'))        {$tbl.ShowFilter        = [bool]$ShowFilter}
        if     ($PSBoundParameters.ContainsKey('ShowFirstColumn'))   {$tbl.ShowFirstColumn   = [bool]$ShowFirstColumn}
        if     ($PSBoundParameters.ContainsKey('ShowLastColumn'))    {$tbl.ShowLastColumn    = [bool]$ShowLastColumn}
        if     ($PSBoundParameters.ContainsKey('ShowRowStripes'))    {$tbl.ShowRowStripes    = [bool]$ShowRowStripes}
        if     ($PSBoundParameters.ContainsKey('ShowColumnStripes')) {$tbl.ShowColumnStripes = [bool]$ShowColumnStripes}
        $tbl.TableStyle = $TableStyle

        if ($PassThru) {return $tbl}
    }
    catch {Write-Warning -Message "Failed adding table '$TableName' to worksheet '$WorksheetName': $_"}
}
