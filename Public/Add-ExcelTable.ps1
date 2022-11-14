function Add-ExcelTable {
    [CmdletBinding()]
    [OutputType([OfficeOpenXml.Table.ExcelTable])]
    param (
        [Parameter(Mandatory=$true)]
        [OfficeOpenXml.ExcelRange]$Range,
        [String]$TableName = "",
        [OfficeOpenXml.Table.TableStyles]$TableStyle = 'Medium6',
        [Switch]$ShowHeader ,
        [Switch]$ShowFilter,
        [Switch]$ShowTotal,
        [hashtable]$TableTotalSettings,
        [Switch]$ShowFirstColumn,
        [Switch]$ShowLastColumn,
        [Switch]$ShowRowStripes,
        [Switch]$ShowColumnStripes,
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
        if     ($PSBoundParameters.ContainsKey('TableTotalSettings') -And $Null -ne $TableTotalSettings)     {
            $tbl.ShowTotal = $true
            foreach ($k in $TableTotalSettings.keys) {
                
                # Get the Function to be added in the totals row
                if ($TableTotalSettings[$k] -is [HashTable]) { 
                    If ($TableTotalSettings[$k].Keys -contains "Function") {
                        $TotalFunction = $TableTotalSettings[$k]["Function"]
                    }
                    Else { Write-Warning -Message "TableTotalSettings parameter for column '$k' needs a key 'Function'" }
                }
                else { 
                    $TotalFunction = [String]($TableTotalSettings[$k]) 
                }
                
                # Add the totals row
                if (-not $tbl.Columns[$k]) {Write-Warning -Message "Table does not have a Column '$k'."}
                elseif ($TotalFunction -match "^=") {
                    ### A function in Excel uses ";" between parameters but the OpenXML parameter separator is ","
                    ### Only replace semicolon when it's NOT somewhere between quotes quotes. 
                    # Get all text between quotes
                    $QuoteMatches = [Regex]::Matches($TotalFunction,"`"[^`"]*`"|'[^']*'") 
                    # Create array with all indexes of characters between quotes (and the quotes themselves)
                    $QuoteCharIndexes = $(
                        Foreach ($QuoteMatch in $QuoteMatches) {
                            (($QuoteMatch.Index)..($QuoteMatch.Index + $QuoteMatch.Length - 1))
                        }
                    )

                    # Get all semicolons
                    $SemiColonMatches = [Regex]::Matches($TotalFunction, ";")
                    # Replace the semicolons of which the index is not in the list of quote-text indexes
                    Foreach ($SemiColonMatch in $SemiColonMatches.Index) {
                        If ($QuoteCharIndexes -notcontains $SemiColonMatch) {
                            $TotalFunction = $TotalFunction.remove($SemiColonMatch,1).Insert($SemiColonMatch,",")
                        }
                    }

                    # Configure the formula. The TotalsRowFunction is automatically set to "Custom" when the TotalsRowFormula is set.
                    $tbl.Columns[$k].TotalsRowFormula = $TotalFunction
                }
                elseif ($TotalFunction -notin @("Average", "Count", "CountNums", "Max", "Min", "None", "StdDev", "Sum", "Var") ) {
                    Write-Warning -Message "'$($TotalFunction)' is not a valid total function."
                }
                else {$tbl.Columns[$k].TotalsRowFunction = $TotalFunction}

                # Set comment on totals row
                If ($TableTotalSettings[$k] -is [HashTable] -and $TableTotalSettings[$k].Keys -contains "Comment" -and ![String]::IsNullOrEmpty($TableTotalSettings[$k]["Comment"])) {
                    $ColumnLetter = [officeOpenXml.ExcelAddress]::GetAddressCol(($tbl.columns | ? { $_.name -eq $k }).Id, $False)
                    $CommentRange = "{0}{1}" -f $ColumnLetter, $tbl.Address.End.Row

                    $CellCommentParams = @{
                        Worksheet    = $tbl.WorkSheet
                        Range        = $CommentRange
                        Text         = $TableTotalSettings[$k]["Comment"]
                    }

                    Set-CellComment  @CellCommentParams
                }
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
