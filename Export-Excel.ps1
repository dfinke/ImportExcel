Function Export-Excel {
    <#
        .SYNOPSIS
            Export data to an Excel worksheet.

        .DESCRIPTION
            Export data to an Excel file and where possible try to convert numbers so Excel recognizes them as numbers instead of text. After all. Excel is a spreadsheet program used for number manipulation and calculations. In case the number conversion is not desired, use the parameter '-NoNumberConversion *'.

        .PARAMETER ConditionalText
            Applies a 'Conditional formatting rule' in Excel on all the cells. When specific conditions are met a rule is triggered.

        .PARAMETER NoNumberConversion
            By default we convert all values to numbers if possible, but this isn't always desirable. NoNumberConversion allows you to add exceptions for the conversion. Wildcards (like '*') are allowed.

        .PARAMETER AutoFilter
            Enables the 'Filter' in Excel on the complete header row. So users can easily sort, filter and/or search the data in the select column from within Excel.

        .PARAMETER AutoSize
            Sizes the width of the Excel column to the maximum width needed to display all the containing data in that cell.

        .PARAMETER Now
            The 'Now' switch is a shortcut that creates automatically a temporary file, enables 'AutoSize', 'AutoFiler' and 'Show', and opens the file immediately.

        .PARAMETER NumberFormat
            Formats all values that can be converted to a number to the format specified.
            
            Examples:
            # integer (not really needed unless you need to round numbers, Excel with use default cell properties)
            '0'

            # integer without displaying the number 0 in the cell
            '#'

            # number with 1 decimal place
            '0.0'

            # number with 2 decimal places
            '0.00'

            # number with 2 decimal places and thousand separator
            '#,##0.00'

            # number with 2 decimal places and thousand separator and money symbol
            '€#,##0.00'

            # percentage (1 = 100%, 0.01 = 1%)
            '0%'

            # Blue color for positive numbers and a red color for negative numbers. All numbers will proceed a dollar sign '$'.
            '[Blue]$#,##0.00;[Red]-$#,##0.00'

        .PARAMETER Show
            Opens the Excel file immediately after creation. Convenient for viewing the results instantly without having to search for the file first.

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
        
        .LINK
            https://github.com/dfinke/ImportExcel
    #>

    [CmdLetBinding()]
    Param(
        $Path,
        [Parameter(ValueFromPipeline=$true)]
        $TargetData,
        [String]$WorkSheetname = 'Sheet1',
        [String]$Title,
        [OfficeOpenXml.Style.ExcelFillStyle]$TitleFillPattern = 'None',
        [Switch]$TitleBold,
        [Int]$TitleSize = 22,
        [System.Drawing.Color]$TitleBackgroundColor,
        [String[]]$PivotRows,
        [String[]]$PivotColumns,
        $PivotData,
        [Switch]$PivotDataToColumn,
        [String]$Password,
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType = 'Pie',
        [Switch]$IncludePivotTable,
        [Switch]$IncludePivotChart,
        [Switch]$NoLegend,
        [Switch]$ShowCategory,
        [Switch]$ShowPercent,
        [Switch]$AutoSize,
        [Switch]$Show,
        [Switch]$NoClobber,
        [Switch]$FreezeTopRow,
        [Switch]$FreezeFirstColumn,
        [Switch]$FreezeTopRowFirstColumn,
        [Int[]]$FreezePane,
        [Switch]$AutoFilter,
        [Switch]$BoldTopRow,
        [Switch]$NoHeader,
        [String]$RangeName,
        [ValidateScript({
            if ($_.Contains(' ')) {
                throw 'Tablename has spaces.'
            }
            elseif (-not $_) {
                throw 'Tablename is null or empty.'
            }
            elseif ($_[0] -notmatch '[a-z]') {
                throw 'Tablename start with invalid character.'
            }
            else {
                $true
            }
        })] 
        [String]$TableName,
        [OfficeOpenXml.Table.TableStyles]$TableStyle = 'Medium6',
        [Object[]]$ExcelChartDefinition,
        [String[]]$HideSheet,
        [Switch]$KillExcel,
        [Switch]$AutoNameRange,
        [Int]$StartRow = 1,
        [Int]$StartColumn = 1,
        [Switch]$PassThru,
        [String]$Numberformat = 'General',
        [String[]]$NoNumberConversion,
        [Object[]]$ConditionalFormat,
        [Object[]]$ConditionalText,
        [ScriptBlock]$CellStyleSB,
        [Switch]$Now
    )

    Begin {
        Function Add-CellValue {
            <#
            .SYNOPSIS
                Save a value in an Excel cell.

            .DESCRIPTION
                DateTime objects are always converted to a DateTime format in Excel. And formulas are always
                saved as formulas.

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

                {$_ -is [DateTime]} {
                    #region Save a date with an international valid format
                    $TargetCell.Value = $_
                    $TargetCell.Style.Numberformat.Format = 'm/d/yy h:mm'
                    Write-Verbose "Cell '$Row`:$ColumnIndex' header '$Name' add value '$_' as date"
                    break
                    #endregion
                }

                {(($NoNumberConversion) -and ($NoNumberConversion -contains $Name)) -or 
                ($NoNumberConversion -eq '*')} {
                    #regioon Save a value without converting to number
                    $TargetCell.Value = $_
                    Write-Verbose "Cell '$Row`:$ColumnIndex' header '$Name' add value '$($TargetCell.Value)' unconverted"
                    break
                    #endregion
                }

                Default {
                    #region Save a value as a number if possible
                    if ($Number = ConvertTo-Number $_) {
                        $TargetCell.Value = $Number
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

        Function Add-Title {
            <#
            .SYNOPSIS
                Add a title row to the Excel worksheet.
            #>

            $ws.Cells[$Row, $StartColumn].Value = $Title
            $ws.Cells[$Row, $StartColumn].Style.Font.Size = $TitleSize

            if ($TitleBold) {
                #set title to Bold if -TitleBold was specified.
                #Otherwise the default will be unbolded.
                $ws.Cells[$Row, $StartColumn].Style.Font.Bold = $True
            }
            $ws.Cells[$Row, $StartColumn].Style.Fill.PatternType = $TitleFillPattern

            #can only set TitleBackgroundColor if TitleFillPattern is something other than None
            if ($TitleBackgroundColor -AND ($TitleFillPattern -ne 'None')) {
                $ws.Cells[$Row, $StartColumn].Style.Fill.BackgroundColor.SetColor($TitleBackgroundColor)
            }
            else {
                Write-Warning "Title Background Color ignored. You must set the TitleFillPattern parameter to a value other than 'None'. Try 'Solid'."
            }
        }

        Function ConvertTo-Number {
            <#
            .SYNOPSIS
                Convert a value to a number
            #>

            Param (
                [String]$Value
            )

            $R = $null

            if ([Double]::TryParse([String]$Value,[System.Globalization.NumberStyles]::Any,
                        [System.Globalization.NumberFormatInfo]::CurrentInfo, [Ref]$R)) {
                $R        
            }
        }

        Function Stop-ExcelProcess {
            <#
            .SYNOPSIS
                Stop the Excel process when it's running.
            #>

            Get-Process excel -ErrorAction Ignore | Stop-Process
            while (Get-Process excel -ErrorAction Ignore) {}
        }

        Try {
            $script:Header = $null

            if ($KillExcel) {
                Stop-ExcelProcess
            }

            if ($Now) {
                $Path = [System.IO.Path]::GetTempFileName() -replace '\.tmp','.xlsx'                
                $Show = $true
                $AutoSize = $true
                $AutoFilter = $true
            }

            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)

            if (Test-Path $Path) {
                Write-Debug "Path '$Path' already exists"
            }

            $pkg = New-Object OfficeOpenXml.ExcelPackage $Path
            $ws  = $pkg | Add-WorkSheet -WorkSheetname $WorkSheetname -NoClobber:$NoClobber

            foreach ($format in $ConditionalFormat ) {
                $target = "Add$($format.Formatter)"
                $rule = ($ws.ConditionalFormatting).PSObject.Methods[$target].Invoke($format.Range, $format.IconType)
                $rule.Reverse = $format.Reverse
            }

            $Row = $StartRow

            if ($Title) {
                Add-Title

                $Row += 1
            }

            $firstTimeThru = $true
            $isDataTypeValueType = $false
            $pattern = 'string|bool|byte|char|decimal|double|float|int|long|sbyte|short|uint|ulong|ushort'
        } 
        Catch {
            if ($AlreadyExists) {
                throw "Failed exporting worksheet '$WorkSheetname' to '$Path': The worksheet '$WorkSheetname' already exists."
            } 
            else {
                throw "Failed exporting worksheet '$WorkSheetname' to '$Path': $_"
            }
        }
    }

    Process {
        Try {
            if ($firstTimeThru) {
                $firstTimeThru = $false
                $isDataTypeValueType = $TargetData.GetType().name -match $pattern
                Write-Debug "DataTypeName is '$($TargetData.GetType().name)' isDataTypeValueType '$isDataTypeValueType'"
            }

            if ($isDataTypeValueType) {
                $ColumnIndex = $StartColumn

                Add-CellValue -TargetCell $ws.Cells[$Row, $ColumnIndex] -CellValue $TargetData

                $ColumnIndex += 1
                $Row += 1
            } 
            else {
                #region Add headers
                if (-not $script:Header) {
                    $ColumnIndex = $StartColumn
                    $script:Header = $TargetData.PSObject.Properties.Name

                    if ($NoHeader) {
                        # Don't push the headers to the spread sheet
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

    End {
        Try {
            if ($AutoNameRange) {
                $totalRows = $ws.Dimension.Rows
                $totalColumns = $ws.Dimension.Columns

                foreach($c in 0..($totalColumns-1)) {
                    $targetRangeName = "$($script:Header[$c])"                

                    $targetColumn = $c+1
                    $theCell = $ws.Cells[2,$targetColumn,$totalRows,$targetColumn ]
                    $ws.Names.Add($targetRangeName, $theCell) | Out-Null

                    if ([OfficeOpenXml.FormulaParsing.ExcelUtilities.ExcelAddressUtil]::IsValidAddress($targetRangeName)) {
                        Write-Warning "AutoNameRange: Property name '$targetRangeName' is also a valid Excel address and may cause issues. Consider renaming the property name."
                    }
                }
            }

            $startAddress=$ws.Dimension.Start.Address
            $dataRange="{0}:{1}" -f $startAddress, $ws.Dimension.End.Address

            Write-Debug "Data Range '$dataRange'"

            if (-not [String]::IsNullOrEmpty($RangeName)) {
                $ws.Names.Add($RangeName, $ws.Cells[$dataRange]) | Out-Null
            }

            if (-not [String]::IsNullOrEmpty($TableName)) {
                $csr = $StartRow
                $csc = $StartColumn
                $cer = $ws.Dimension.End.Row
                $cec = $script:Header.Count

                $targetRange = $ws.Cells[$csr, $csc, $cer,$cec]
                $tbl = $ws.Tables.Add($targetRange, $TableName)
                $tbl.TableStyle = $TableStyle
            }

            if ($IncludePivotTable) {
                $pivotTableName = $WorkSheetname + 'PivotTable'
                $wsPivot = $pkg | Add-WorkSheet -WorkSheetname $pivotTableName -NoClobber:$NoClobber

                $wsPivot.View.TabSelected = $true

                $pivotTableDataName=$WorkSheetname + 'PivotTableData'

                if ($Title) {
                    $startAddress = 'A2'
                }

                $pivotTable = $wsPivot.PivotTables.Add($wsPivot.Cells['A1'], $ws.Cells[$dataRange], $pivotTableDataName)

                if ($PivotRows) {
                    foreach ($Row in $PivotRows) {
                        $null = $pivotTable.RowFields.Add($pivotTable.Fields[$Row])
                    }
                }

                if ($PivotColumns) {
                    foreach ($Column in $PivotColumns) {
                        $null = $pivotTable.ColumnFields.Add($pivotTable.Fields[$Column])
                    }
                }

                if ($PivotData) {
                    if ($PivotData -is [HashTable] -or $PivotData -is [System.Collections.Specialized.OrderedDictionary]) {
                        $PivotData.Keys | ForEach-Object {
                            $df = $pivotTable.DataFields.Add($pivotTable.Fields[$_])
                            $df.Function = $PivotData.$_
                        }
                    }
                    else {
                        foreach ($Item in $PivotData) {
                            $df = $pivotTable.DataFields.Add($pivotTable.Fields[$Item])
                            $df.Function = 'Count'
                        }
                    }
                    if ($PivotDataToColumn) {
                        $pivotTable.DataOnRows = $false
                    }
                }

                if ($IncludePivotChart) {
                    $chart = $wsPivot.Drawings.AddChart('PivotChart', $ChartType, $pivotTable)
                    $chart.DataLabel.ShowCategory = $ShowCategory
                    $chart.DataLabel.ShowPercent = $ShowPercent

                    if ($NoLegend) {
                        $chart.Legend.Remove()
                    }

                    $chart.SetPosition(1, 0, 6, 0)
                    $chart.SetSize(600, 400)
                }
            }

            if ($Password) {
                $ws.Protection.SetPassword($Password)
            }

            if ($AutoFilter) {
                $ws.Cells[$dataRange].AutoFilter=$true
            }

            if ($FreezeTopRow) {
                $ws.View.FreezePanes(2,1)
            }

            if ($FreezeTopRowFirstColumn) {
                $ws.View.FreezePanes(2,2)
            }

            if ($FreezeFirstColumn) {
                $ws.View.FreezePanes(1,2)
            }

            if ($FreezePane) {
                $freezeRow,$freezeColumn=$FreezePane
                if (-not $freezeColumn -or $freezeColumn -eq 0) {
                    $freezeColumn=1
                }

                if ($freezeRow -gt 1) {
                    $ws.View.FreezePanes($freezeRow,$freezeColumn)
                }
            }
            if ($BoldTopRow) {
                $range=$ws.Dimension.Address -replace $ws.Dimension.Rows, '1'
                $ws.Cells[$range].Style.Font.Bold = $true
            }
            if ($AutoSize) {
                $ws.Cells.AutoFitColumns()
            }

            foreach ($Sheet in $HideSheet) {
                $pkg.Workbook.WorkSheets[$Sheet].Hidden = 'Hidden'
            }

            $chartCount=0
            foreach ($chartDef in $ExcelChartDefinition) {
                $ChartName = 'Chart' + (Split-Path -Leaf ([System.IO.path]::GetTempFileName())) -replace 'tmp|\.',''
                $chart = $ws.Drawings.AddChart($ChartName, $chartDef.ChartType)
                $chart.Title.Text = $chartDef.Title

                if ($chartDef.NoLegend) {
                    $chart.Legend.Remove()
                }
            
                if ($chart.Datalabel -ne $null) {
                    $chart.Datalabel.ShowCategory  = $chartDef.ShowCategory
                    $chart.Datalabel.ShowPercent   = $chartDef.ShowPercent
                }

                $chart.SetPosition($chartDef.Row, $chartDef.RowOffsetPixels,$chartDef.Column, $chartDef.ColumnOffsetPixels)
                $chart.SetSize($chartDef.Width, $chartDef.Height)

                $chartDefCount = @($chartDef.YRange).Count
                if ($chartDefCount -eq 1) {
                    $Series=$chart.Series.Add($chartDef.YRange, $chartDef.XRange)
                
                    $SeriesHeader=$chartDef.SeriesHeader
                    if (-not $SeriesHeader) {
                        $SeriesHeader = 'Series 1'
                    }

                    $Series.Header = $SeriesHeader
                } else {
                    for($idx = 0; $idx -lt $chartDefCount; $idx += 1) {
                        $Series=$chart.Series.Add($chartDef.YRange[$idx], $chartDef.XRange)                    

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

                    $Range=$targetConditionalText.Range
                    if (-not $Range) {
                        $Range = $ws.Dimension.Address
                    }

                    $rule=($ws.Cells[$Range].ConditionalFormatting).PSObject.Methods[$target].Invoke()

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
                $pkg.Save()
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
