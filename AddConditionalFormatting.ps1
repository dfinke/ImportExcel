Function Add-ConditionalFormatting {
    <#
      .Synopsis
        Adds conditional formatting to worksheet.
      .Description
        Conditional formatting allows excel to
        * Mark cells with Icons depending on their value
        * Show a databar whose length indicates the value or a 2 or 3 color scale where the color indicate the relative value
        * Change the color, font, or number format of cells which meet given criteria
        Add-ConditionalFormatting allows these to be set; for fine tuning of the rules you can use the -PassThru switch,
        which will return the rule so that you can modify things which are specific to that type of rule,
        for example the values which correspond to each icon in an Icon set.
      .Example
        >
        PS> $excel = $avdata | Export-Excel -Path (Join-path $FilePath "\Machines.XLSX" ) -WorksheetName "Server Anti-Virus" -AutoSize -FreezeTopRow -AutoFilter -PassThru
        Add-ConditionalFormatting -WorkSheet $excel.Workbook.Worksheets[1] -Address "b2:b1048576" -ForeGroundColor "RED"     -RuleType ContainsText -ConditionValue "2003"
        Add-ConditionalFormatting -WorkSheet $excel.Workbook.Worksheets[1] -Address "i2:i1048576" -ForeGroundColor "RED"     -RuleType ContainsText -ConditionValue "Disabled"
        $excel.Workbook.Worksheets[1].Cells["D1:G1048576"].Style.Numberformat.Format = [cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern
        $excel.Workbook.Worksheets[1].Row(1).style.font.bold = $true
        $excel.Save() ; $excel.Dispose()

        Here Export-Excel is called with the -passThru parameter so the Excel Package object is stored in $Excel
        The desired worksheet is selected and the then columns" B" and "I" are conditionally formatted (excluding the top row) to show red text if
        the columns contain "2003" or "Disabled respectively. A fixed date format is then applied to columns D..G, and the top row is formatted.
        Finally the workbook is saved and the Excel object closed.
      .Example
        >
        >PS  $r = Add-ConditionalFormatting -WorkSheet $excel.Workbook.Worksheets[1] -Range "B1:B100" -ThreeIconsSet Flags -Passthru
        $r.Reverse = $true ;   $r.Icon1.Type = "Num"; $r.Icon2.Type = "Num" ; $r.Icon2.value = 100 ; $r.Icon3.type = "Num" ;$r.Icon3.value = 1000

        Again Export-Excel has been called with -passthru leaving a package object in $Excel
        This time B1:B100 has been conditionally formatted with 3 icons, using the flags icon set.
        Add-ConditionalFormatting does not provide access to every option in the formatting rule, so passthru has been used and the
        rule is modified to apply the flags in reverse order, and boundaries for the number which will set the split are set to 100 and 1000
      .Example
        Add-ConditionalFormatting -WorkSheet $sheet -Range "D2:D1048576" -DataBarColor Red

        This time $sheet holds an ExcelWorkseet object and databars are add to all of column D except for the tip row.
      .Example
        Add-ConditionalFormatting -Address $worksheet.cells["FinishPosition"] -RuleType Equal    -ConditionValue 1 -ForeGroundColor Purple -Bold -Priority 1 -StopIfTrue

        In this example a named range is used to select the cells where the formula should apply. If a cell in the "FinishPosition" range is 1, then the text is turned to Bold & Purple.
        This rule is moved to first in the priority list, and where cells have a value of 1, no other rules will be processed.
    #>
    Param (
        #A block of cells to format - you can use a named range with -Address $ws.names[1] or  $ws.cells["RangeName"]
        [Parameter(Mandatory = $true, Position = 0)]
        [Alias("Range")]
        $Address ,
        #The worksheet where the format is to be applied
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet ,
        #A standard named-rule - Top / Bottom / Less than / Greater than / Contains etc.
        [Parameter(Mandatory = $true, ParameterSetName = "NamedRule", Position = 1)]
        [OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType]$RuleType ,
        #Text colour for matching objects
        [Parameter(ParameterSetName = "NamedRule")]
        [Alias("ForeGroundColour")]
        [System.Drawing.Color]$ForeGroundColor,
        #Colour for databar type charts
        [Parameter(Mandatory = $true, ParameterSetName = "DataBar")]
        [Alias("DataBarColour")]
        [System.Drawing.Color]$DataBarColor,
        #One of the three-icon set types (e.g. Traffic Lights)
        [Parameter(Mandatory = $true, ParameterSetName = "ThreeIconSet")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting3IconsSetType]$ThreeIconsSet,
        #A four-icon set name
        [Parameter(Mandatory = $true, ParameterSetName = "FourIconSet")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting4IconsSetType]$FourIconsSet,
        #A five-icon set name
        [Parameter(Mandatory = $true, ParameterSetName = "FiveIconSet")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting5IconsSetType]$FiveIconsSet,
        #Use the icon set in reverse order, or reverse the orders of Two- & Three-Color Scales
        [Parameter(ParameterSetName = "NamedRule")]
        [Parameter(ParameterSetName = "ThreeIconSet")]
        [Parameter(ParameterSetName = "FourIconSet")]
        [Parameter(ParameterSetName = "FiveIconSet")]
        [switch]$Reverse,
        #A value for the condition (for example 2000 if the test is 'lessthan 2000'; Formulas should begin with "=" )
        [Parameter(ParameterSetName = "NamedRule")]
        $ConditionValue,
        #A second value for the conditions like "between x and Y"
        [Parameter(ParameterSetName = "NamedRule")]
        $ConditionValue2,
        #Background colour for matching items
        [Parameter(ParameterSetName = "NamedRule")]
        [System.Drawing.Color]$BackgroundColor,
        #Background pattern for matching items
        [Parameter(ParameterSetName = "NamedRule")]
        [OfficeOpenXml.Style.ExcelFillStyle]$BackgroundPattern = [OfficeOpenXml.Style.ExcelFillStyle]::None ,
        #Secondary colour when a background pattern requires it
        [Parameter(ParameterSetName = "NamedRule")]
        [System.Drawing.Color]$PatternColor,
        #Sets the numeric format for matching items
        [Parameter(ParameterSetName = "NamedRule")]
        $NumberFormat,
        #Put matching items in bold face
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$Bold,
        #Put matching items in italic
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$Italic,
        #Underline matching items
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$Underline,
        #Strikethrough text of matching items
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$StrikeThru,
        #Prevent the processing of subsequent rules
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$StopIfTrue,
        #Set the sequence for rule processing
        [int]$Priority,
        #If specified pass the rule back to the caller to allow additional customization.
        [switch]$Passthru
    )

    #Allow conditional formatting to work like Set-ExcelRange (with single ADDRESS parameter), split it to get worksheet and range of cells.
    If ($Address -is [OfficeOpenXml.Table.ExcelTable]) {
            $WorkSheet = $Address.Address.Worksheet
            $Address   = $Address.Address.Address
    }
    elseif  ($Address.Address -and $Address.Worksheet -and -not $WorkSheet) { #Address is a rangebase or similar
        $WorkSheet = $Address.Worksheet[0]
        $Address   = $Address.Address
    }
    elseif ($Address -is [String] -and $WorkSheet -and $WorkSheet.Names[$Address] ) { #Address is the name of a named range.
        $Address = $WorkSheet.Names[$Address].Address
    }
    if (($Address -is [OfficeOpenXml.ExcelRow]    -and -not $WorkSheet) -or
        ($Address -is [OfficeOpenXml.ExcelColumn] -and -not $WorkSheet) ){  #EPPLUs Can't get the worksheet object from a row or column object, so bail if that was tried
        Write-Warning -Message "Add-ConditionalFormatting does not support Row or Column objects as an address; use a worksheet and/or specify 'R:R' or 'C:C' instead. "; return
    }
    elseif ($Address -is [OfficeOpenXml.ExcelRow]) {  #But if we have a column or row object and a worksheet (I don't know *why*) turn them into a string for the range
            $Address = "$($Address.Row):$($Address.Row)"
    }
    elseif ($Address -is [OfficeOpenXml.ExcelColumn]) {
        $Address = [OfficeOpenXml.ExcelAddress]::new(1,$address.ColumnMin,1,$address.ColumnMax).Address -replace '1',''
        if ($Address -notmatch ':') {$Address = "$Address`:$Address"}
    }
    if ( $Address -is [string] -and $Address -match "!") {$Address = $Address -replace '^.*!',''}
    #By this point we should have a worksheet object whose ConditionalFormatting collection we will add to. If not, bail.
    if (-not $worksheet -or $WorkSheet -isnot [OfficeOpenXml.ExcelWorksheet]) {write-warning "You need to provide a worksheet object." ; return}
    #region create a rule of the right type
    if     ($RuleType -match 'IconSet$') {Write-warning -Message "You cannot configure a IconSet rule in this way; please use -$RuleType <SetName>." ; return}
    if     ($PSBoundParameters.ContainsKey("ThreeIconsSet" )      ) {$rule =  $WorkSheet.ConditionalFormatting.AddThreeIconSet($Address , $ThreeIconsSet)}
    elseif ($PSBoundParameters.ContainsKey("FourIconsSet"  )      ) {$rule =  $WorkSheet.ConditionalFormatting.AddFourIconSet( $Address , $FourIconsSet )}
    elseif ($PSBoundParameters.ContainsKey("FiveIconsSet"  )      ) {$rule =  $WorkSheet.ConditionalFormatting.AddFiveIconSet( $Address , $FiveIconsSet )}
    elseif ($PSBoundParameters.ContainsKey("DataBarColor"  )      ) {$rule =  $WorkSheet.ConditionalFormatting.AddDatabar(     $Address , $DataBarColor )}
    else                                                            {$rule = ($WorkSheet.ConditionalFormatting)."Add$RuleType"($Address )                }
    if     ($Reverse)  {
            if     ($rule.type -match 'IconSet$'   )                {$rule.reverse = $true}
            elseif ($rule.type -match 'ColorScale$')                {$temp =$rule.LowValue.Color ; $rule.LowValue.Color = $rule.HighValue.Color; $rule.HighValue.Color = $temp}
            else   {Write-Warning -Message "-Reverse was ignored because $ruletype does not support it."}
    }
    #endregion
    #region set the rule conditions
    #for lessThan/GreaterThan/Equal/Between conditions make sure that strings are wrapped in quotes. Formulas should be passed with = which will be stripped.
    if     ($RuleType -match "Than|Equal|Between" ) {
        if ($ConditionValue) {
                $number = $Null
                #if the condition type is not a value type, but parses as a number, make it the number
                if ($ConditionValue -isnot [System.ValueType] -and [Double]::TryParse($ConditionValue, [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::CurrentInfo, [Ref]$number) ) {
                         $ConditionValue  = $number
                } #else if it is not a value type, or a formula, or wrapped in quotes, wrap it in quotes.
                elseif (($ConditionValue -isnot [System.ValueType])-and ($ConditionValue  -notmatch '^=') -and ($ConditionValue  -notmatch '^".*"$') ) {
                         $ConditionValue  = '"' + $ConditionValue +'"'
                }
        }
        if ($ConditionValue2) {
                $number = $Null
                if ($ConditionValue -isnot [System.ValueType] -and [Double]::TryParse($ConditionValue2, [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::CurrentInfo, [Ref]$number) ) {
                         $ConditionValue2 = $number
                }
                elseif (($ConditionValue -isnot [System.ValueType]) -and ($ConditionValue2 -notmatch '^=') -and ($ConditionValue2 -notmatch '^".*"$') ) {
                         $ConditionValue2  = '"' + $ConditionValue2 + '"'
                }
        }
    }
    #But we don't usually want quotes round containstext | beginswith type rules. Can't be Certain they need to be removed, so warn the user their condition might be wrong
    if     ($RuleType -match "Text|With" -and $ConditionValue -match '^".*"$'  ) {
            Write-Warning -Message "The condition will look for the quotes at the start and end."
    }
    if     ($PSBoundParameters.ContainsKey("ConditionValue" ) -and
            $RuleType -match "Top|Botom"                          ) {$rule.Rank      = $ConditionValue }
    if     ($PSBoundParameters.ContainsKey("ConditionValue" ) -and
            $RuleType -match "StdDev"                             ) {$rule.StdDev    = $ConditionValue }
    if     ($PSBoundParameters.ContainsKey("ConditionValue" ) -and
            $RuleType -match "Than|Equal|Expression"              ) {$rule.Formula   = ($ConditionValue  -replace '^=','') }
    if     ($PSBoundParameters.ContainsKey("ConditionValue" ) -and
            $RuleType -match "Text|With"                          ) {$rule.Text      = ($ConditionValue  -replace '^=','') }
    if     ($PSBoundParameters.ContainsKey("ConditionValue" ) -and
            $PSBoundParameters.ContainsKey("ConditionValue2") -and
            $RuleType -match "Between"                            ) {
                                                                     $rule.Formula   = ($ConditionValue  -replace '^=','');
                                                                     $rule.Formula2  = ($ConditionValue2 -replace '^=','')
    }
    if     ($PSBoundParameters.ContainsKey("StopIfTrue")          ) {$rule.StopIfTrue = $StopIfTrue }
    if     ($PSBoundParameters.ContainsKey("Priority")            ) {$rule.Priority   = $Priority }
    #endregion
    #region set the rule format
    if     ($PSBoundParameters.ContainsKey("NumberFormat"     )   ) {$rule.Style.NumberFormat.Format        = (Expand-NumberFormat  $NumberFormat)             }
    if     ($Underline                                            ) {$rule.Style.Font.Underline             = [OfficeOpenXml.Style.ExcelUnderLineType]::Single }
    elseif ($PSBoundParameters.ContainsKey("Underline"        )   ) {$rule.Style.Font.Underline             = [OfficeOpenXml.Style.ExcelUnderLineType]::None   }
    if     ($PSBoundParameters.ContainsKey("Bold"             )   ) {$rule.Style.Font.Bold                  = [boolean]$Bold       }
    if     ($PSBoundParameters.ContainsKey("Italic"           )   ) {$rule.Style.Font.Italic                = [boolean]$Italic     }
    if     ($PSBoundParameters.ContainsKey("StrikeThru"       )   ) {$rule.Style.Font.Strike                = [boolean]$StrikeThru }
    if     ($PSBoundParameters.ContainsKey("ForeGroundColor"  )   ) {$rule.Style.Font.Color.color           = $ForeGroundColor     }
    if     ($PSBoundParameters.ContainsKey("BackgroundColor"  )   ) {$rule.Style.Fill.BackgroundColor.color = $BackgroundColor     }
    if     ($PSBoundParameters.ContainsKey("BackgroundPattern")   ) {$rule.Style.Fill.PatternType           = $BackgroundPattern   }
    if     ($PSBoundParameters.ContainsKey("PatternColor"     )   ) {$rule.Style.Fill.PatternColor.color    = $PatternColor        }
    #endregion
    #Allow further tweaking by returning the rule, if passthru specified
    if     ($Passthru)  {$rule}
}