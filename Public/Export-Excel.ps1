function Export-Excel {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [OutputType([OfficeOpenXml.ExcelPackage])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "")]
    param(
        [Parameter(ParameterSetName = 'Default', Position = 0)]
        [String]$Path,
        [Parameter(Mandatory = $true, ParameterSetName = "Package")]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        [Parameter(ValueFromPipeline = $true)]
        [Alias('TargetData')]
        $InputObject,
        [Switch]$Calculate,
        [Switch]$Show,
        [String]$WorksheetName = 'Sheet1',
        [Alias("PW")]
        [String]$Password,
        [switch]$ClearSheet,
        [switch]$Append,
        [String]$Title,
        [OfficeOpenXml.Style.ExcelFillStyle]$TitleFillPattern = 'Solid',
        [Switch]$TitleBold,
        [Int]$TitleSize = 22,
        $TitleBackgroundColor,
        [parameter(DontShow = $true)]
        [Switch]$IncludePivotTable,
        [String]$PivotTableName,
        [String[]]$PivotRows,
        [String[]]$PivotColumns,
        $PivotData,
        [String[]]$PivotFilter,
        [Switch]$PivotDataToColumn,
        [Hashtable]$PivotTableDefinition,
        [Switch]$IncludePivotChart,
        [Alias('ChartType')]
        [OfficeOpenXml.Drawing.Chart.eChartType]$PivotChartType = 'Pie',
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
                if (-not $_) { throw 'RangeName is null or empty.' }
                elseif ($_[0] -notmatch '[a-z]') { throw 'RangeName starts with an invalid character.' }
                else { $true }
            })]
        [String]$RangeName,
        [Alias('Table')]
        $TableName,
        [OfficeOpenXml.Table.TableStyles]$TableStyle = [OfficeOpenXml.Table.TableStyles]::Medium6,
        [HashTable]$TableTotalSettings,
        [Switch]$BarChart,
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
        [alias('PT')]
        [Switch]$PassThru,
        [String]$Numberformat = 'General',
        [string[]]$ExcludeProperty,
        [Switch]$NoAliasOrScriptPropeties,
        [Switch]$DisplayPropertySet,
        [String[]]$NoNumberConversion,
        [String[]]$NoHyperLinkConversion,
        [Object[]]$ConditionalFormat,
        [Object[]]$ConditionalText,
        [Object[]]$Style,
        [ScriptBlock]$CellStyleSB,
        #If there is already content in the workbook the sheet with the PivotTable will not be active UNLESS Activate is specified
        [switch]$Activate,
        [Parameter(ParameterSetName = 'Default')]
        [Switch]$Now,
        [Switch]$ReturnRange,
        #By default PivotTables have Totals for each Row (on the right) and for each column at the bottom. This allows just one or neither to be selected.
        [ValidateSet("Both", "Columns", "Rows", "None")]
        [String]$PivotTotals = "Both",
        #Included for compatibility - equivalent to -PivotTotals "None"
        [Switch]$NoTotalsInPivot,
        [Switch]$ReZip
    )

    begin {
        $numberRegex = [Regex]'\d'
        $isDataTypeValueType = $false
        if ($NoClobber) { Write-Warning -Message "-NoClobber parameter is no longer used" }
        #Open the file, get the worksheet, and decide where in the sheet we are writing, and if there is a number format to apply.
        try {
            $script:Header = $null
            if ($Append -and $ClearSheet) { throw "You can't use -Append AND -ClearSheet." ; return }
            #To force -Now not to format as a table, allow $false in -TableName to be "No table"
            $TableName = if ($null -eq $TableName -or ($TableName -is [bool] -and $false -eq $TableName)) { $null } else { [String]$TableName }
            if ($Now -or (-not $Path -and -not $ExcelPackage) ) {
                if (-not $PSBoundParameters.ContainsKey("Path")) { $Path = [System.IO.Path]::GetTempFileName() -replace '\.tmp', '.xlsx' }
                if (-not $PSBoundParameters.ContainsKey("Show")) { $Show = $true }
                if (-not $PSBoundParameters.ContainsKey("AutoSize")) { $AutoSize = $true }
                #"Now" option will create a table, unless something passed in TableName/Table Style. False in TableName will block autocreation
                if (-not $PSBoundParameters.ContainsKey("TableName") -and
                    -not $PSBoundParameters.ContainsKey("TableStyle") -and
                    -not $AutoFilter) {
                    $TableName = '' # later rely on distinction between NULL and ""
                }
            }
            if ($ExcelPackage) {
                $pkg = $ExcelPackage
                $Path = $pkg.File
            }
            Else { $pkg = Open-ExcelPackage -Path $Path -Create -KillExcel:$KillExcel -Password:$Password }
        }
        catch { throw "Could not open Excel Package $path" }
        try {
            $params = @{WorksheetName = $WorksheetName }
            foreach ($p in @("ClearSheet", "MoveToStart", "MoveToEnd", "MoveBefore", "MoveAfter", "Activate")) { if ($PSBoundParameters[$p]) { $params[$p] = $PSBoundParameters[$p] } }
            $ws = $pkg | Add-Worksheet @params
            if ($ws.Name -ne $WorksheetName) {
                Write-Warning -Message "The Worksheet name has been changed from $WorksheetName to $($ws.Name), this may cause errors later."
                $WorksheetName = $ws.Name
            }
        }
        catch { throw "Could not get worksheet $WorksheetName" }
        try {
            if ($Append -and $ws.Dimension) {
                #if there is a title or anything else above the header row, append needs to be combined wih a suitable startrow parameter
                $headerRange = $ws.Dimension.Address -replace "\d+$", $StartRow
                #using a slightly odd syntax otherwise header ends up as a 2D array
                $ws.Cells[$headerRange].Value | ForEach-Object -Begin { $Script:header = @() } -Process { $Script:header += $_ }
                $NoHeader = $true
                #if we did not get AutoNameRange, but headers have ranges of the same name make autoNameRange True, otherwise make it false
                if (-not $AutoNameRange) {
                    $AutoNameRange = $true ; foreach ($h in $header) { if ($ws.names.name -notcontains $h) { $AutoNameRange = $false } }
                }
                #if we did not get a Rangename but there is a Range which covers the active part of the sheet, set Rangename to that.
                if (-not $RangeName -and $ws.names.where({ $_.name[0] -match '[a-z]' })) {
                    $theRange = $ws.names.where({
                         ($_.Name[0] -match '[a-z]' ) -and
                         ($_.Start.Row -eq $StartRow) -and
                         ($_.Start.Column -eq $StartColumn) -and
                         ($_.End.Row -eq $ws.Dimension.End.Row) -and
                         ($_.End.Column -eq $ws.Dimension.End.column) } , 'First', 1)
                    if ($theRange) { $rangename = $theRange.name }
                }

                #if we did not get a table name but there is a table which covers the active part of the sheet, set table name to that, and don't do anything with autofilter
                $existingTable = $ws.Tables.Where({ $_.address.address -eq $ws.dimension.address }, 'First', 1)
                if ($null -eq $TableName -and $existingTable) {
                    $TableName = $existingTable.Name
                    $TableStyle = $existingTable.StyleName -replace "^TableStyle", ""
                    $AutoFilter = $false
                }
                #if we did not get $autofilter but a filter range is set and it covers the right area, set autofilter to true
                elseif (-not $AutoFilter -and $ws.Names['_xlnm._FilterDatabase']) {
                    if ( ($ws.Names['_xlnm._FilterDatabase'].Start.Row -eq $StartRow) -and
                         ($ws.Names['_xlnm._FilterDatabase'].Start.Column -eq $StartColumn) -and
                         ($ws.Names['_xlnm._FilterDatabase'].End.Row -eq $ws.Dimension.End.Row) -and
                         ($ws.Names['_xlnm._FilterDatabase'].End.Column -eq $ws.Dimension.End.Column) ) { $AutoFilter = $true }
                }

                $row = $ws.Dimension.End.Row
                Write-Debug -Message ("Appending: headers are " + ($script:Header -join ", ") + " Start row is $row")
                if ($Title) { Write-Warning -Message "-Title Parameter is ignored when appending." }
            }
            elseif ($Title) {
                #Can only add a title if not appending!
                $row = $StartRow
                $ws.Cells[$row, $StartColumn].Value = $Title
                $ws.Cells[$row, $StartColumn].Style.Font.Size = $TitleSize

                if ($PSBoundParameters.ContainsKey("TitleBold")) {
                    #Set title to Bold face font if -TitleBold was specified.
                    #Otherwise the default will be unbolded.
                    $ws.Cells[$row, $StartColumn].Style.Font.Bold = [boolean]$TitleBold
                }
                if ($TitleBackgroundColor ) {
                    if ($TitleBackgroundColor -is [string]) { $TitleBackgroundColor = [System.Drawing.Color]::$TitleBackgroundColor }
                    $ws.Cells[$row, $StartColumn].Style.Fill.PatternType = $TitleFillPattern
                    $ws.Cells[$row, $StartColumn].Style.Fill.BackgroundColor.SetColor($TitleBackgroundColor)
                }
                $row ++ ; $startRow ++
            }
            else { $row = $StartRow }
            $ColumnIndex = $StartColumn
            $Numberformat = Expand-NumberFormat -NumberFormat $Numberformat
            if ((-not $ws.Dimension) -and ($Numberformat -ne $ws.Cells.Style.Numberformat.Format)) {
                $ws.Cells.Style.Numberformat.Format = $Numberformat
                $setNumformat = $false
            }
            else { $setNumformat = ($Numberformat -ne $ws.Cells.Style.Numberformat.Format) }
        }
        catch { throw "Failed preparing to export to worksheet '$WorksheetName' to '$Path': $_" }
        #region Special case -inputobject passed a dataTable object
        <# If inputObject was passed via the pipeline it won't be visible until the process block, we will only see it here if it was passed as a parameter
          if it is a data table don't do foreach on it (slow) - put the whole table in and set dates on date columns,
          set things up for the end block, and skip the process block #>
        if ($InputObject -is [System.Data.DataTable]) {
            if ($Append -and $ws.dimension) {
                $row ++
                $null = $ws.Cells[$row, $StartColumn].LoadFromDataTable($InputObject, $false )
                if ($TableName -or $PSBoundParameters.ContainsKey('TableStyle')) {
                    Add-ExcelTable -Range $ws.Cells[$ws.Dimension] -TableName $TableName -TableStyle $TableStyle -TableTotalSettings $TableTotalSettings
                }
            }
            else {
                #Change TableName if $TableName is non-empty; don't leave caller with a renamed table!
                $orginalTableName = $InputObject.TableName
                if ($PSBoundParameters.ContainsKey("TableName")) {
                    $InputObject.TableName = $TableName
                }
                while ($InputObject.TableName -in $pkg.Workbook.Worksheets.Tables.name) {
                    Write-Warning "Table name $($InputObject.TableName) is not unique, adding '_' to it "
                    $InputObject.TableName += "_"
                }
                #Insert as a table, if Tablestyle didn't arrive as a default, or $TableName non-null - even if empty
                if ($null -ne $TableName -or $PSBoundParameters.ContainsKey("TableStyle")) {
                    $null = $ws.Cells[$row, $StartColumn].LoadFromDataTable($InputObject, (-not $noHeader), $TableStyle )
                    # Workaround for EPPlus not marking the empty row on an empty table as dummy row.
                    if ($InputObject.Rows.Count -eq 0) {
                        ($ws.Tables | Select-Object -Last 1).TableXml.table.SetAttribute('insertRow', 1)
                    }
                }
                else {
                    $null = $ws.Cells[$row, $StartColumn].LoadFromDataTable($InputObject, (-not $noHeader) )
                }
                $InputObject.TableName = $orginalTableName
            }
            foreach ($c in $InputObject.Columns.where({ $_.datatype -eq [datetime] })) {
                Set-ExcelColumn -Worksheet $ws -Column ($c.Ordinal + $StartColumn) -NumberFormat 'Date-Time'
            }
            foreach ($c in $InputObject.Columns.where({ $_.datatype -eq [timespan] })) {
                Set-ExcelColumn -Worksheet $ws -Column ($c.Ordinal + $StartColumn) -NumberFormat '[h]:mm:ss'
            }
            $ColumnIndex += $InputObject.Columns.Count - 1
            if ($noHeader) { $row += $InputObject.Rows.Count - 1 }
            else { $row += $InputObject.Rows.Count }
            $null = $PSBoundParameters.Remove('InputObject')
            $firstTimeThru = $false
        }
        #endregion
        else { $firstTimeThru = $true }
    }

    process {
        if ($PSBoundParameters.ContainsKey("InputObject")) {
            try {
                if ($null -eq $InputObject) { $row += 1 }
                foreach ($TargetData in $InputObject) {
                    if ($firstTimeThru) {
                        $firstTimeThru = $false
                        $isDataTypeValueType = ($null -eq $TargetData) -or ($TargetData.GetType().name -match 'string|timespan|datetime|bool|byte|char|decimal|double|float|int|long|sbyte|short|uint|ulong|ushort|URI|ExcelHyperLink')
                        if ($isDataTypeValueType ) {
                            $script:Header = @(".")       # dummy value to make sure we go through the "for each name in $header"
                            if (-not $Append) { $row -= 1 } # By default row will be 1, it is incremented before inserting values (so it ends pointing at final row.);  si first data row is 2 - move back up 1 if there is no header .
                        }
                        if ($null -ne $TargetData) { Write-Debug "DataTypeName is '$($TargetData.GetType().name)' isDataTypeValueType '$isDataTypeValueType'" }
                    }
                    #region Add headers - if we are appending, or we have been through here once already we will have the headers
                    if (-not $script:Header) {
                        if ($DisplayPropertySet -and $TargetData.psStandardmembers.DefaultDisplayPropertySet.ReferencedPropertyNames) {
                            $script:Header = $TargetData.psStandardmembers.DefaultDisplayPropertySet.ReferencedPropertyNames.Where( { $_ -notin $ExcludeProperty })
                        }
                        else {
                            if ($NoAliasOrScriptPropeties) { $propType = "Property" } else { $propType = "*" }
                            $script:Header = $TargetData.PSObject.Properties.where( { $_.MemberType -like $propType }).Name
                        }
                        foreach ($exclusion in $ExcludeProperty) { $script:Header = $script:Header -notlike $exclusion }
                        if ($NoHeader) {
                            # Don't push the headers to the spreadsheet
                            $row -= 1
                        }
                        else {
                            $ColumnIndex = $StartColumn
                            foreach ($Name in $script:Header) {
                                $ws.Cells[$row, $ColumnIndex].Value = $Name
                                Write-Verbose "Cell '$row`:$ColumnIndex' add header '$Name'"
                                $ColumnIndex += 1
                            }
                        }
                    }
                    #endregion
                    #region Add non header values
                    $row += 1
                    $ColumnIndex = $StartColumn
                    <#
                 For each item in the header OR for the Data item if this is a simple Type or data table :
                   If it is a date insert with one of Excel's built in formats - recognized as "Date and time to be localized"
                   if it is a timespan insert with a built in format for elapsed hours, minutes and seconds
                   if its  any other numeric insert as is , setting format if need be.
                   Preserve URI, Insert a data table, convert non string objects to string.
                   For strings, check for fomula, URI or Number, before inserting as a string  (ignore nulls) #>
                    foreach ($Name in $script:Header) {
                        if ($isDataTypeValueType) { $v = $TargetData }
                        else { $v = $TargetData.$Name }
                        try {
                            if ($v -is [DateTime]) {
                                $ws.Cells[$row, $ColumnIndex].Value = $v
                                $ws.Cells[$row, $ColumnIndex].Style.Numberformat.Format = 'm/d/yy h:mm' # This is not a custom format, but a preset recognized as date and localized.
                            }
                            elseif ($v -is [TimeSpan]) {
                                $ws.Cells[$row, $ColumnIndex].Value = $v
                                $ws.Cells[$row, $ColumnIndex].Style.Numberformat.Format = '[h]:mm:ss'
                            }
                            elseif ($v -is [System.ValueType]) {
                                $ws.Cells[$row, $ColumnIndex].Value = $v
                                if ($setNumformat) { $ws.Cells[$row, $ColumnIndex].Style.Numberformat.Format = $Numberformat }
                            }
                            elseif ($v -is [uri] -and
                                    $NoHyperLinkConversion -ne '*' -and
                                    $NoHyperLinkConversion -notcontains $Name ) {
                                $ws.Cells[$row, $ColumnIndex].HyperLink = $v
                                $ws.Cells[$row, $ColumnIndex].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
                                $ws.Cells[$row, $ColumnIndex].Style.Font.UnderLine = $true
                            }
                            elseif ($v -isnot [String] ) {
                                #Other objects or null.
                                if ($null -ne $v) { $ws.Cells[$row, $ColumnIndex].Value = $v.toString() }
                            }
                            elseif ($v[0] -eq '=') {
                                $ws.Cells[$row, $ColumnIndex].Formula = ($v -replace '^=', '')
                                if ($setNumformat) { $ws.Cells[$row, $ColumnIndex].Style.Numberformat.Format = $Numberformat }
                            }
                            elseif ( $NoHyperLinkConversion -ne '*' -and # Put the check for 'NoHyperLinkConversion is null' first to skip checking for wellformedstring
                                    $NoHyperLinkConversion -notcontains $Name -and
                                    [System.Uri]::IsWellFormedUriString($v , [System.UriKind]::Absolute)
                                ) { 
                                if ($v -match "^xl://internal/") {
                                    $referenceAddress = $v -replace "^xl://internal/" , ""
                                    $display = $referenceAddress -replace "!A1$"   , ""
                                    $h = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $referenceAddress , $display
                                    $ws.Cells[$row, $ColumnIndex].HyperLink = $h
                                }
                                else { $ws.Cells[$row, $ColumnIndex].HyperLink = $v }   #$ws.Cells[$row, $ColumnIndex].Value = $v.AbsoluteUri
                                $ws.Cells[$row, $ColumnIndex].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
                                $ws.Cells[$row, $ColumnIndex].Style.Font.UnderLine = $true
                            }
                            else {
                                $number = $null
                                if ( $NoNumberConversion -ne '*' -and # Check if NoNumberConversion isn't specified. Put this first as it's going to stop the if clause. Quicker than putting regex check first
                                    $numberRegex.IsMatch($v) -and # and if it contains digit(s) - this syntax is quicker than -match for many items and cuts out slow checks for non numbers
                                    $NoNumberConversion -notcontains $Name -and
                                    [Double]::TryParse($v, [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::CurrentInfo, [Ref]$number)
                                ) {
                                    $ws.Cells[$row, $ColumnIndex].Value = $number
                                    if ($setNumformat) { $ws.Cells[$row, $ColumnIndex].Style.Numberformat.Format = $Numberformat }
                                }
                                else {
                                    $ws.Cells[$row, $ColumnIndex].Value = $v
                                }
                            }
                        }
                        catch { Write-Warning -Message "Could not insert the '$Name' property at Row $row, Column $ColumnIndex" }
                        $ColumnIndex += 1
                    }
                    $ColumnIndex -= 1 # column index will be the last column whether isDataTypeValueType was true or false
                    #endregion
                }
            }
            catch { throw "Failed exporting data to worksheet '$WorksheetName' to '$Path': $_" }
        }
    }

    end {
        if ($firstTimeThru -and $ws.Dimension) {
            $LastRow = $ws.Dimension.End.Row
            $LastCol = $ws.Dimension.End.Column
            $endAddress = $ws.Dimension.End.Address
        }
        else {
            $LastRow = $row
            $LastCol = $ColumnIndex
            $endAddress = [OfficeOpenXml.ExcelAddress]::GetAddress($LastRow , $LastCol)
        }
        $startAddress = [OfficeOpenXml.ExcelAddress]::GetAddress($StartRow, $StartColumn)
        $dataRange = "{0}:{1}" -f $startAddress, $endAddress

        Write-Debug "Data Range '$dataRange'"
        if ($AutoNameRange) {
            try {
                if (-not $script:header) {
                    # if there aren't any headers, use the the first row of data to name the ranges: this is the last point that headers will be used.
                    $headerRange = $ws.Dimension.Address -replace "\d+$", $StartRow
                    #using a slightly odd syntax otherwise header ends up as a 2D array
                    $ws.Cells[$headerRange].Value | ForEach-Object -Begin { $Script:header = @() } -Process { $Script:header += $_ }
                    if ($PSBoundParameters.ContainsKey($TargetData)) {
                        #if Export was called with data that writes no header start the range at $startRow ($startRow is data)
                        $targetRow = $StartRow
                    }
                    else { $targetRow = $StartRow + 1 }                   #if Export was called without data to add names (assume $startRow is a header) or...
                }                                                         #          ... called with data that writes a header, then start the range at $startRow + 1
                else { $targetRow = $StartRow + 1 }

                #Dimension.start.row always seems to be one so we work out the target row
                #, but start.column is the first populated one and .Columns is the count of populated ones.
                # if we have 5 columns from 3 to 8, headers are numbered 0..4, so that is in the for loop and used for getting the name...
                # but we have to add the start column on when referencing positions
                foreach ($c in 0..($LastCol - $StartColumn)) {
                    $targetRangeName = @($script:Header)[$c]  #Let Add-ExcelName fix (and warn about) bad names
                    Add-ExcelName  -RangeName $targetRangeName -Range $ws.Cells[$targetRow, ($StartColumn + $c ), $LastRow, ($StartColumn + $c )]
                    try {
                        #this test can throw with some names, surpress any error
                        if ([OfficeOpenXml.FormulaParsing.ExcelUtilities.ExcelAddressUtil]::IsValidAddress(($targetRangeName -replace '\W' , '_' ))) {
                            Write-Warning -Message "AutoNameRange: Property name '$targetRangeName' is also a valid Excel address and may cause issues. Consider renaming the property."
                        }
                    }
                    catch {
                        Write-Warning -Message "AutoNameRange: Testing '$targetRangeName' caused an error. This should be harmless, but a change of property name may be needed.."
                    }
                }
            }
            catch { Write-Warning -Message "Failed adding named ranges to worksheet '$WorksheetName': $_" }
        }
        #Empty string is not allowed as a name for ranges or tables.
        if ($RangeName) { Add-ExcelName  -Range $ws.Cells[$dataRange] -RangeName $RangeName }

        #Allow table to be inserted by specifying Name, or Style or both; only process autoFilter if there is no table (they clash).
        if ($null -ne $TableName -or $PSBoundParameters.ContainsKey('TableStyle')) {
            #Already inserted Excel table if input was a DataTable
            if ($InputObject -isnot [System.Data.DataTable]) {
                Add-ExcelTable -Range $ws.Cells[$dataRange] -TableName $TableName -TableStyle $TableStyle -TableTotalSettings $TableTotalSettings
            }
        }
        elseif ($AutoFilter) {
            try {
                $ws.Cells[$dataRange].AutoFilter = $true
                Write-Verbose -Message "Enabled autofilter. "
            }
            catch { Write-Warning -Message "Failed adding autofilter to worksheet '$WorksheetName': $_" }
        }

        if ($PivotTableDefinition) {
            foreach ($item in $PivotTableDefinition.GetEnumerator()) {
                $params = $item.value
                if ($Activate) { $params.Activate = $true }
                if ($params.keys -notcontains 'SourceRange' -and
                   ($params.Keys -notcontains 'SourceWorksheet' -or $params.SourceWorksheet -eq $WorksheetName)) { $params.SourceRange = $dataRange }
                if ($params.Keys -notcontains 'SourceWorksheet') { $params.SourceWorksheet = $ws }
                if ($params.Keys -notcontains 'NoTotalsInPivot' -and $NoTotalsInPivot  ) { $params.PivotTotals = 'None' }
                if ($params.Keys -notcontains 'PivotTotals' -and $PivotTotals      ) { $params.PivotTotals = $PivotTotals }
                if ($params.Keys -notcontains 'PivotDataToColumn' -and $PivotDataToColumn) { $params.PivotDataToColumn = $true }

                Add-PivotTable -ExcelPackage $pkg -PivotTableName $item.key @Params
            }
        }
        if ($IncludePivotTable -or $IncludePivotChart -or $PivotData) {
            $params = @{
                'SourceRange' = $dataRange
            }
            if ($PivotTableName -and ($pkg.workbook.worksheets.tables.name -contains $PivotTableName)) {
                Write-Warning -Message "The selected PivotTable name '$PivotTableName' is already used as a table name. Adding a suffix of 'Pivot'."
                $PivotTableName += 'Pivot'
            }

            if ($PivotTableName) { $params.PivotTableName = $PivotTableName }
            else { $params.PivotTableName = $WorksheetName + 'PivotTable' }
            if ($Activate) { $params.Activate = $true }
            if ($PivotFilter) { $params.PivotFilter = $PivotFilter }
            if ($PivotRows) { $params.PivotRows = $PivotRows }
            if ($PivotColumns) { $Params.PivotColumns = $PivotColumns }
            if ($PivotData) { $Params.PivotData = $PivotData }
            if ($NoTotalsInPivot) { $params.PivotTotals = "None" }
            Elseif ($PivotTotals) { $params.PivotTotals = $PivotTotals }
            if ($PivotDataToColumn) { $params.PivotDataToColumn = $true }
            if ($IncludePivotChart -or
                $PSBoundParameters.ContainsKey('PivotChartType')) {
                $params.IncludePivotChart = $true
                $Params.ChartType = $PivotChartType
                if ($ShowCategory) { $params.ShowCategory = $true }
                if ($ShowPercent) { $params.ShowPercent = $true }
                if ($NoLegend) { $params.NoLegend = $true }
            }
            Add-PivotTable -ExcelPackage $pkg -SourceWorksheet $ws   @params
        }

        try {
            #Allow single switch or two seperate ones.
            if ($FreezeTopRowFirstColumn -or ($FreezeTopRow -and $FreezeFirstColumn)) {
                if ($Title) {
                    $ws.View.FreezePanes(3, 2)
                    Write-Verbose -Message "Froze title and header rows and first column"
                }
                else {
                    $ws.View.FreezePanes(2, 2)
                    Write-Verbose -Message "Froze top row and first column"
                }
            }
            elseif ($FreezeTopRow) {
                if ($Title) {
                    $ws.View.FreezePanes(3, 1)
                    Write-Verbose -Message "Froze title and header rows"
                }
                else {
                    $ws.View.FreezePanes(2, 1)
                    Write-Verbose -Message "Froze top row"
                }
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
        catch { Write-Warning -Message "Failed adding Freezing the panes in worksheet '$WorksheetName': $_" }

        if ($PSBoundParameters.ContainsKey("BoldTopRow")) {
            #it sets bold as far as there are populated cells: for whole row could do $ws.row($x).style.font.bold = $true
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
            catch { Write-Warning -Message "Failed setting the top row to bold in worksheet '$WorksheetName': $_" }
        }
        if ($AutoSize -and -not $env:NoAutoSize) {
            try {
                #Don't fit the all the columns in the sheet; if we are adding cells beside things with hidden columns, that unhides them
                if ($MaxAutoSizeRows -and $MaxAutoSizeRows -lt $LastRow ) {
                    $AutosizeRange = [OfficeOpenXml.ExcelAddress]::GetAddress($startRow, $StartColumn, $MaxAutoSizeRows , $LastCol)
                    $ws.Cells[$AutosizeRange].AutoFitColumns()
                }
                else { $ws.Cells[$dataRange].AutoFitColumns() }
                Write-Verbose -Message "Auto-sized columns"
            }
            catch { Write-Warning -Message "Failed autosizing columns of worksheet '$WorksheetName': $_" }
        }
        elseif ($AutoSize) { Write-Warning -Message "Auto-fitting columns is not available with this OS configuration." }

        foreach ($Sheet in $HideSheet) {
            try {
                $pkg.Workbook.Worksheets.Where({ $_.Name -like $sheet }) | ForEach-Object {
                    $_.Hidden = 'Hidden'
                    Write-verbose -Message "Sheet '$($_.Name)' Hidden."
                }
            }
            catch { Write-Warning -Message  "Failed hiding worksheet '$sheet': $_" }
        }
        foreach ($Sheet in $UnHideSheet) {
            try {
                $pkg.Workbook.Worksheets.Where({ $_.Name -like $sheet }) | ForEach-Object {
                    $_.Hidden = 'Visible'
                    Write-verbose -Message "Sheet '$($_.Name)' shown"
                }
            }
            catch { Write-Warning -Message  "Failed showing worksheet '$sheet': $_" }
        }
        if (-not $pkg.Workbook.Worksheets.Where({ $_.Hidden -eq 'visible' })) {
            Write-Verbose -Message "No Sheets were left visible, making $WorksheetName visible"
            $ws.Hidden = 'Visible'
        }

        foreach ($chartDef in $ExcelChartDefinition) {
            if ($chartDef -is [System.Management.Automation.PSCustomObject]) {
                $params = @{}
                $chartDef.PSObject.Properties | ForEach-Object { if ( $null -ne $_.value) { $params[$_.name] = $_.value } }
                Add-ExcelChart -Worksheet $ws @params
            }
            elseif ($chartDef -is [hashtable] -or $chartDef -is [System.Collections.Specialized.OrderedDictionary]) {
                Add-ExcelChart -Worksheet $ws @chartDef
            }
        }

        if ($Calculate) {
            try { [OfficeOpenXml.CalculationExtension]::Calculate($ws) }
            catch { Write-Warning "One or more errors occured while calculating, save will continue, but there may be errors in the workbook. $_" }
        }

        if ($Barchart -or $PieChart -or $LineChart -or $ColumnChart) {
            if ($NoHeader) { $FirstDataRow = $startRow }
            else { $FirstDataRow = $startRow + 1 }
            $range = [OfficeOpenXml.ExcelAddress]::GetAddress($FirstDataRow, $startColumn, $FirstDataRow, $lastCol )
            $xCol = $ws.cells[$range] | Where-Object { $_.value -is [string] } | ForEach-Object { $_.start.column } | Sort-Object | Select-Object -first 1
            if (-not $xcol) {
                $xcol = $StartColumn
                $range = [OfficeOpenXml.ExcelAddress]::GetAddress($FirstDataRow, ($startColumn + 1), $FirstDataRow, $lastCol )
            }
            $yCol = $ws.cells[$range] | Where-Object { $_.value -is [valueType] -or $_.Formula } | ForEach-Object { $_.start.column } | Sort-Object | Select-Object -first 1
            if (-not ($xCol -and $ycol)) { Write-Warning -Message "Can't identify a string column and a number column to use as chart labels and data. " }
            else {
                $params = @{
                    XRange = [OfficeOpenXml.ExcelAddress]::GetAddress($FirstDataRow, $xcol , $lastrow, $xcol)
                    YRange = [OfficeOpenXml.ExcelAddress]::GetAddress($FirstDataRow, $ycol , $lastrow, $ycol)
                    Title  = ''
                    Column = ($lastCol + 1)
                    Width  = 800
                }
                if ($ShowPercent) { $params["ShowPercent"] = $true }
                if ($ShowCategory) { $params["ShowCategory"] = $true }
                if ($NoLegend) { $params["NoLegend"] = $true }
                if (-not $NoHeader) { $params["SeriesHeader"] = $ws.Cells[$startRow, $YCol].Value }
                if ($ColumnChart) { $Params["chartType"] = "ColumnStacked" }
                elseif ($Barchart) { $Params["chartType"] = "BarStacked" }
                elseif ($PieChart) { $Params["chartType"] = "PieExploded3D" }
                elseif ($LineChart) { $Params["chartType"] = "Line" }

                Add-ExcelChart -Worksheet $ws @params
            }
        }

        # It now doesn't matter if the conditional formating rules are passed in $conditionalText or $conditional format.
        # Just one with an alias for compatiblity it will break things for people who are using both at once
        foreach ($c in  (@() + $ConditionalText + $ConditionalFormat) ) {
            try {
                #we can take an object with a .ConditionalType property made by New-ConditionalText or with a .Formatter Property made by New-ConditionalFormattingIconSet or a hash table
                if ($c.ConditionalType) {
                    $cfParams = @{RuleType = $c.ConditionalType; ConditionValue = $c.Text ;
                        BackgroundColor = $c.BackgroundColor; BackgroundPattern = $c.PatternType  ;
                        ForeGroundColor = $c.ConditionalTextColor
                    }
                    if ($c.Range) { $cfParams.Range = $c.Range }
                    else { $cfParams.Range = $ws.Dimension.Address }
                    Add-ConditionalFormatting -Worksheet $ws @cfParams
                    Write-Verbose -Message "Added conditional formatting to range $($c.range)"
                }
                elseif ($c.formatter) {
                    switch ($c.formatter) {
                        "ThreeIconSet" { Add-ConditionalFormatting -Worksheet $ws -ThreeIconsSet $c.IconType -range $c.range -reverse:$c.reverse -ShowIconOnly:$c.ShowIconOnly}
                        "FourIconSet" { Add-ConditionalFormatting -Worksheet $ws  -FourIconsSet $c.IconType -range $c.range -reverse:$c.reverse -ShowIconOnly:$c.ShowIconOnly}
                        "FiveIconSet" { Add-ConditionalFormatting -Worksheet $ws  -FiveIconsSet $c.IconType -range $c.range -reverse:$c.reverse -ShowIconOnly:$c.ShowIconOnly}
                    }
                    Write-Verbose -Message "Added conditional formatting to range $($c.range)"
                }
                elseif ($c -is [hashtable] -or $c -is [System.Collections.Specialized.OrderedDictionary]) {
                    if (-not $c.Range -or $c.Address) { $c.Address = $ws.Dimension.Address }
                    Add-ConditionalFormatting -Worksheet $ws @c
                }
            }
            catch { throw "Error applying conditional formatting to worksheet $_" }
        }
        foreach ($s in $Style) {
            if (-not $s.Range) { $s["Range"] = $ws.Dimension.Address }
            Set-ExcelRange -Worksheet $ws @s
        }
        if ($CellStyleSB) {
            try {
                $TotalRows = $ws.Dimension.Rows
                $LastColumn = $ws.Dimension.Address -replace "^.*:(\w*)\d+$" , '$1'
                & $CellStyleSB $ws $TotalRows $LastColumn
            }
            catch { Write-Warning -Message "Failed processing CellStyleSB in worksheet '$WorksheetName': $_" }
        }

        #Can only add password, may want to support -password $Null removing password.
        if ($Password) {
            try {
                $ws.Protection.SetPassword($Password)
                Write-Verbose -Message 'Set password on workbook'
            }
            catch { throw "Failed setting password for worksheet '$WorksheetName': $_" }
        }

        if ($PassThru) { $pkg }
        else {
            if ($ReturnRange) { $dataRange }

            if ($Password) { $pkg.Save($Password) }
            else { $pkg.Save() }
            Write-Verbose -Message "Saved workbook $($pkg.File)"
            if ($ReZip) {
                Invoke-ExcelReZipFile -ExcelPackage $pkg
            }

            $pkg.Dispose()

            if ($Show) { Invoke-Item $Path }
        }
    }
}
