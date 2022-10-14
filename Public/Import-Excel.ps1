function Import-Excel {
    [CmdLetBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSPossibleIncorrectUsageOfAssignmentOperator', '', Justification = 'Intentional')]
    param (
        [Alias('FullName')]
        [Parameter(ParameterSetName = "PathA", Mandatory, ValueFromPipelineByPropertyName, ValueFromPipeline, Position = 0 )]
        [Parameter(ParameterSetName = "PathB", Mandatory, ValueFromPipelineByPropertyName, ValueFromPipeline, Position = 0 )]
        [Parameter(ParameterSetName = "PathC", Mandatory, ValueFromPipelineByPropertyName, ValueFromPipeline, Position = 0 )]
        [String[]]$Path,
        [Parameter(ParameterSetName = "PackageA", Mandatory)]
        [Parameter(ParameterSetName = "PackageB", Mandatory)]
        [Parameter(ParameterSetName = "PackageC", Mandatory)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        [Alias('Sheet')]
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [String[]]$WorksheetName,
        [Parameter(ParameterSetName = 'PathB'   , Mandatory)]
        [Parameter(ParameterSetName = 'PackageB', Mandatory)]
        [String[]]$HeaderName ,
        [Parameter(ParameterSetName = 'PathC'   , Mandatory)]
        [Parameter(ParameterSetName = 'PackageC', Mandatory)]
        [Switch]$NoHeader     ,
        [Alias('HeaderRow', 'TopRow')]
        [ValidateRange(1, 1048576)]
        [Int]$StartRow = 1,
        [Alias('StopRow', 'BottomRow')]
        [Int]$EndRow ,
        [Alias('LeftColumn')]
        [Int]$StartColumn = 1,
        [Alias('RightColumn')]
        [Int]$EndColumn  ,
        [Switch]$DataOnly,
        [string[]]$AsText,
        [string[]]$AsDate,
        [ValidateNotNullOrEmpty()]
        [String]$Password,
        [Int[]]$ImportColumns,
        [Switch]$Raw
    )
    end {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        if ($input) {
            $Paths = $input
        }
        elseif ($Path) {
            $Paths = $Path
        }
        else {
            $Paths = ''
        }
        function Get-PropertyNames {
            <#
            .SYNOPSIS
                Create objects containing the column number and the column name for each of the different header types.
            #>
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "Name would be incorrect, and command is not exported")]
            param(
                [Parameter(Mandatory)]
                [Int[]]$Columns,
                [Parameter(Mandatory)]
                [Int]$StartRow
            )

            try {
                if ($ImportColumns) {
                    $end = $sheet.Dimension.End.Column
                    # Check $ImportColumns
                    if ($ImportColumns[0] -le 0) { throw "The first entry in ImportColumns must be equal or greater 1" ; return }
                    # Check $StartColumn and $EndColumn
                    if (($StartColumn -ne 1) -or ($EndColumn -ne $end)) { Write-Warning -Message "If ImportColumns is set, then individual StartColumn and EndColumn will be ignored." }
                    # Replace $Columns with $ImportColumns
                    $Columns = $ImportColumns
                }

                if ($HeaderName) {
                    $i = 0
                    foreach ($H in $HeaderName) {
                        $H | Select-Object @{N = 'Column'; E = { $Columns[$i] } }, @{N = 'Value'; E = { $H } }
                        $i++
                    }
                }
                elseif ($NoHeader) {
                    $i = 0
                    foreach ($C in $Columns) {
                        $i++
                        $C | Select-Object @{N = 'Column'; E = { $_ } }, @{N = 'Value'; E = { 'P' + $i } }
                    }
                }

                else {
                    if ($StartRow -lt 1) {
                        throw 'The top row can never be less than 1 when we need to retrieve headers from the worksheet.' ; return
                    }

                    foreach ($C in $Columns) {
                        #allow "False" or "0" to be column headings
                        $sheet.Cells[$StartRow, $C] | Where-Object { -not [string]::IsNullOrEmpty($_.Value) } | Select-Object @{N = 'Column'; E = { $C } }, Value
                    }
                }
            }
            catch {
                throw "Failed creating property names: $_" ; return
            }
        }
        foreach ($Path in $Paths) {
            if ($path) {
                $extension = [System.IO.Path]::GetExtension($Path)
                if ($extension -notmatch '.xlsx$|.xlsm$') {
                    throw "Import-Excel does not support reading this extension type $($extension)"
                }

                $resolvedPath = (Resolve-Path $Path -ErrorAction SilentlyContinue)
                if ($resolvedPath) {
                    $Path = $resolvedPath.ProviderPath
                }
                else {
                    throw "'$($Path)' file not found"
                }

                $stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Path, 'Open', 'Read', 'ReadWrite'
                $ExcelPackage = New-Object -TypeName OfficeOpenXml.ExcelPackage
                if ($Password) { $ExcelPackage.Load($stream, $Password) }
                else { $ExcelPackage.Load($stream) }
            }
            try {
                #Select worksheet
                if ($WorksheetName -eq '*') { $Worksheet = $ExcelPackage.Workbook.Worksheets }
                elseif (-not  $WorksheetName) { $Worksheet = $ExcelPackage.Workbook.Worksheets[1] }
                elseif (-not ($Worksheet = $ExcelPackage.Workbook.Worksheets[$WorksheetName])) {
                    throw "Worksheet '$WorksheetName' not found, the workbook only contains the worksheets '$($ExcelPackage.Workbook.Worksheets)'. If you only wish to select the first worksheet, please remove the '-WorksheetName' parameter." ; return
                }

                $xlBook = [Ordered]@{}
                foreach ($sheet in $Worksheet) {
                    if ($Worksheet.Count -gt 1 -or $Paths.Count -gt 1) {
                        <#
                            Needed under these conditions to handle sheets of different number of Row/Col 
                                - When reading more than one xlsx file
                                - When reading more than one worksheet in the same file
                        #>
                        $EndRow = $null
                        $EndColumn = $null
                    }

                    $targetSheetname = $sheet.Name
                    $xlBook["$targetSheetname"] = @()
                    #region Get rows and columns
                    #If we are doing dataonly it is quicker to work out which rows to ignore before processing the cells.
                    if (-not $EndRow   ) { $EndRow = $sheet.Dimension.End.Row }
                    if (-not $EndColumn) { $EndColumn = $sheet.Dimension.End.Column }
                    $endAddress = [OfficeOpenXml.ExcelAddress]::TranslateFromR1C1("R[$EndRow]C[$EndColumn]", 0, 0)
                    if ($DataOnly) {
                        #If we are using headers startrow will be the header-row so examine data from startRow + 1,
                        if ($NoHeader) { $range = "A" + ($StartRow     ) + ":" + $endAddress }
                        else { $range = "A" + ($StartRow + 1 ) + ":" + $endAddress }
                        #We're going to look at every cell and build 2 hash tables holding rows & columns which contain data.
                        #Want to Avoid 'select unique' operations & large Sorts, becuse time time taken increases with square
                        #of number of items (PS uses heapsort at large size). Instead keep a list of what we have seen,
                        #using Hash tables: "we've seen it" is all we need, no need to worry about "seen it before" / "Seen it many times".
                        $colHash = @{ }
                        $rowHash = @{ }
                        foreach ($cell in $sheet.Cells[$range]) {
                            if ($null -ne $cell.Value ) { $colHash[$cell.Start.Column] = 1; $rowHash[$cell.Start.row] = 1 }
                        }
                        $rows = (   $StartRow..$EndRow   ).Where( { $rowHash[$_] })
                        $columns = ($StartColumn..$EndColumn).Where( { $colHash[$_] })
                    }
                    else {
                        $Columns = $StartColumn .. $EndColumn  ; if ($StartColumn -gt $EndColumn) { Write-Warning -Message "Selecting columns $StartColumn to $EndColumn might give odd results." }
                        if ($NoHeader) { $rows = $StartRow..$EndRow ; if ($StartRow -gt $EndRow) { Write-Warning -Message "Selecting rows $StartRow to $EndRow might give odd results." } }
                        elseif ($HeaderName) { $rows = $StartRow..$EndRow }
                        else {
                            $rows = (1 + $StartRow)..$EndRow
                            if ($StartRow -eq 1 -and $EndRow -eq 1) {
                                $rows = 0
                            }
                        }

                        # ; if ($StartRow -ge $EndRow) { Write-Warning -Message "Selecting $StartRow as the header with data in $(1+$StartRow) to $EndRow might give odd results." } }
                    }
                    #endregion
                    #region Create property names
                    if ((-not $Columns) -or (-not ($PropertyNames = Get-PropertyNames -Columns $Columns -StartRow $StartRow))) {
                        throw "No column headers found on top row '$StartRow'. If column headers in the worksheet are not a requirement then please use the '-NoHeader' or '-HeaderName' parameter."; return
                    }
                    if ($Duplicates = $PropertyNames | Group-Object Value | Where-Object Count -GE 2) {
                        throw "Duplicate column headers found on row '$StartRow' in columns '$($Duplicates.Group.Column)'. Column headers must be unique, if this is not a requirement please use the '-NoHeader' or '-HeaderName' parameter."; return
                    }
                    #endregion
                    if (-not $rows) {
                        Write-Warning "Worksheet '$WorksheetName' in workbook '$Path' contains no data in the rows after top row '$StartRow'"
                    }
                    else {
                        #region Create one object per row
                        if ($AsText -or $AsDate) {
                            <#join items in AsText together with ~~~ . Escape any regex special characters...
                        # which turns "*" into "\*" make it ".*". Convert ~~~ to $|^ and top and tail with ^%;
                        So if we get "Week", "[Time]" and "*date*" ; make the expression ^week$|^\[Time\]$|^.*Date.*$
                        $make a regex for this which is case insensitive (option 1) and compiled (option 8)
                        #>
                            $TextColExpression = ''
                            if ($AsText) {
                                $TextColExpression += '(?<astext>^' + [regex]::Escape($AsText -join '~~~').replace('\*', '.*').replace('~~~', '$|^') + '$)'
                            }
                            if ($AsText -and $AsDate) {
                                $TextColExpression += "|"
                            }
                            if ($AsDate) {
                                $TextColExpression += '(?<asDate>^' + [regex]::Escape($AsDate -join '~~~').replace('\*', '.*').replace('~~~', '$|^') + '$)'
                            }
                            $TextColRegEx = New-Object -TypeName regex -ArgumentList $TextColExpression , 9
                        }
                        else { $TextColRegEx = $null }
                        foreach ($R in $rows) {
                            #Disabled write-verbose for speed
                            #  Write-Verbose "Import row '$R'"
                            $NewRow = [Ordered]@{ }
                            if ($TextColRegEx) {
                                foreach ($P in $PropertyNames) {
                                    $MatchTest = $TextColRegEx.Match($P.value)
                                    if ($MatchTest.groups.name -eq "astext") {
                                        $NewRow[$P.Value] = $sheet.Cells[$R, $P.Column].Text
                                    }
                                    elseif ($MatchTest.groups.name -eq "asdate" -and $sheet.Cells[$R, $P.Column].Value -is [System.ValueType]) {
                                        $NewRow[$P.Value] = [datetime]::FromOADate(($sheet.Cells[$R, $P.Column].Value))
                                    }
                                    else { $NewRow[$P.Value] = $sheet.Cells[$R, $P.Column].Value }
                                }
                            }
                            else {
                                foreach ($P in $PropertyNames) {
                                    $NewRow[$P.Value] = $sheet.Cells[$R, $P.Column].Value
                                    #    Write-Verbose "Import cell '$($Worksheet.Cells[$R, $P.Column].Address)' with property name '$($p.Value)' and value '$($Worksheet.Cells[$R, $P.Column].Value)'."
                                }
                            }
                            $xlBook["$targetSheetname"] += [PSCustomObject]$NewRow
                        }
                        #endregion
                    }
                }
            }
            catch { throw "Failed importing the Excel workbook '$Path' with worksheet '$WorksheetName': $_"; return }
            finally {
                # $EndRow = 0
                # $EndColumn = 0
                if ($Path) { $stream.close(); $ExcelPackage.Dispose() }

                if ($Raw) {
                    foreach ($entry in $xlbook.GetEnumerator()) {
                        $entry.Value
                    }
                }
                elseif ($Worksheet.Count -eq 1) {
                    $xlBook["$targetSheetname"]
                }
                else {
                    $xlBook
                }
            }
        }
    }
}
