function Add-ConditionalFormatting {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [Alias("Range")]
        $Address ,
        [OfficeOpenXml.ExcelWorksheet]$Worksheet ,
        [Parameter(Mandatory = $true, ParameterSetName = "NamedRule", Position = 1)]
        [OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType]$RuleType ,
        [Parameter(ParameterSetName = "NamedRule")]
        [Alias("ForegroundColour","FontColor")]
        $ForegroundColor,
        [Parameter(Mandatory = $true, ParameterSetName = "DataBar")]
        [Alias("DataBarColour")]
        $DataBarColor,
        [Parameter(Mandatory = $true, ParameterSetName = "ThreeIconSet")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting3IconsSetType]$ThreeIconsSet,
        [Parameter(Mandatory = $true, ParameterSetName = "FourIconSet")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting4IconsSetType]$FourIconsSet,
        [Parameter(Mandatory = $true, ParameterSetName = "FiveIconSet")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting5IconsSetType]$FiveIconsSet,
        [Parameter(ParameterSetName = "NamedRule")]
        [Parameter(ParameterSetName = "ThreeIconSet")]
        [Parameter(ParameterSetName = "FourIconSet")]
        [Parameter(ParameterSetName = "FiveIconSet")]
        [switch]$Reverse,
        [switch]$ShowIconOnly,
        [Parameter(ParameterSetName = "NamedRule",Position = 2)]
        $ConditionValue,
        [Parameter(ParameterSetName = "NamedRule",Position = 3)]
        $ConditionValue2,
        [Parameter(ParameterSetName = "NamedRule")]
        $BackgroundColor,
        [Parameter(ParameterSetName = "NamedRule")]
        [OfficeOpenXml.Style.ExcelFillStyle]$BackgroundPattern = [OfficeOpenXml.Style.ExcelFillStyle]::None ,
        [Parameter(ParameterSetName = "NamedRule")]
        $PatternColor,
        [Parameter(ParameterSetName = "NamedRule")]
        $NumberFormat,
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$Bold,
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$Italic,
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$Underline,
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$StrikeThru,
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$StopIfTrue,
        [int]$Priority,
        [switch]$PassThru
    )

    #Allow conditional formatting to work like Set-ExcelRange (with single ADDRESS parameter), split it to get worksheet and range of cells.
    if ($Address -is [OfficeOpenXml.Table.ExcelTable]) {
            $Worksheet = $Address.Address.Worksheet
            $Address   = $Address.Address.Address
    }
    elseif  ($Address.Address -and $Address.Worksheet -and -not $Worksheet) { #Address is a rangebase or similar
        $Worksheet = $Address.Worksheet[0]
        $Address   = $Address.Address
    }
    elseif ($Address -is [String] -and $Worksheet -and $Worksheet.Names[$Address] ) { #Address is the name of a named range.
        $Address = $Worksheet.Names[$Address].Address
    }
    if (($Address -is [OfficeOpenXml.ExcelRow]    -and -not $Worksheet) -or
        ($Address -is [OfficeOpenXml.ExcelColumn] -and -not $Worksheet) ){  #EPPLUs Can't get the worksheet object from a row or column object, so bail if that was tried
        Write-Warning -Message "Add-ConditionalFormatting does not support Row or Column objects as an address; use a worksheet and/or specify 'R:R' or 'C:C' instead. "; return
    }
    elseif ($Address -is [OfficeOpenXml.ExcelRow]) {  #But if we have a column or row object and a worksheet (I don't know *why*) turn them into a string for the range
            $Address = "$($Address.Row):$($Address.Row)"
    }
    elseif ($Address -is [OfficeOpenXml.ExcelColumn]) {
        $Address = (New-Object 'OfficeOpenXml.ExcelAddress' @(1, $address.ColumnMin, 1, $address.ColumnMax)).Address -replace '1',''
        if ($Address -notmatch ':') {$Address = "$Address`:$Address"}
    }
    if ( $Address -is [string] -and $Address -match "!") {$Address = $Address -replace '^.*!',''}
    #By this point we should have a worksheet object whose ConditionalFormatting collection we will add to. If not, bail.
    if (-not $worksheet -or $Worksheet -isnot [OfficeOpenXml.ExcelWorksheet]) {write-warning "You need to provide a worksheet object." ; return}
    #region create a rule of the right type
    if     ($RuleType -match 'IconSet$') {Write-warning -Message "You cannot configure a Icon-Set rule in this way; please use -$RuleType <SetName>." ; return}
    if ($PSBoundParameters.ContainsKey("DataBarColor"  )      ) {if ($DataBarColor -is [string]) {$DataBarColor = [System.Drawing.Color]::$DataBarColor }
                                                                     $rule =  $Worksheet.ConditionalFormatting.AddDatabar(     $Address , $DataBarColor )
    }
    elseif ($PSBoundParameters.ContainsKey("ThreeIconsSet" )      ) {$rule =  $Worksheet.ConditionalFormatting.AddThreeIconSet($Address , $ThreeIconsSet)}
    elseif ($PSBoundParameters.ContainsKey("FourIconsSet"  )      ) {$rule =  $Worksheet.ConditionalFormatting.AddFourIconSet( $Address , $FourIconsSet )}
    elseif ($PSBoundParameters.ContainsKey("FiveIconsSet"  )      ) {$rule =  $Worksheet.ConditionalFormatting.AddFiveIconSet( $Address , $FiveIconsSet )}
    else                                                            {$rule = ($Worksheet.ConditionalFormatting)."Add$RuleType"($Address )                }
    If($ShowIconOnly) {
        $rule.ShowValue = $false
    }
    if     ($Reverse)  {
            if     ($rule.type -match 'IconSet$'   )                {$rule.reverse = $true}
            elseif ($rule.type -match 'ColorScale$')                {$temp =$rule.LowValue.Color ; $rule.LowValue.Color = $rule.HighValue.Color; $rule.HighValue.Color = $temp}
            else   {Write-Warning -Message "-Reverse was ignored because $RuleType does not support it."}
    }
    #endregion
    #region set the rule conditions
    #for lessThan/GreaterThan/Equal/Between conditions make sure that strings are wrapped in quotes. Formulas should be passed with = which will be stripped.
    if     ($RuleType -match "Than|Equal|Between" ) {
        if  ($PSBoundParameters.ContainsKey("ConditionValue" )) {
                $number = $Null
                #if the condition type is not a value type, but parses as a number, make it the number
                if ($ConditionValue -isnot [System.ValueType] -and [Double]::TryParse($ConditionValue, [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::CurrentInfo, [Ref]$number) ) {
                         $ConditionValue  = $number
                } #else if it is not a value type, or a formula, or wrapped in quotes, wrap it in quotes.
                elseif (($ConditionValue -isnot [System.ValueType])-and ($ConditionValue  -notmatch '^=') -and ($ConditionValue  -notmatch '^".*"$') ) {
                         $ConditionValue  = '"' + $ConditionValue +'"'
                }
        }
        if  ($PSBoundParameters.ContainsKey("ConditionValue2")) {
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
            $RuleType -match "Top|Bottom"                          ) {$rule.Rank      = $ConditionValue }
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
    if     ($PSBoundParameters.ContainsKey("ForeGroundColor"  )   ) {if ($ForeGroundColor -is [string])      {$ForeGroundColor = [System.Drawing.Color]::$ForeGroundColor }
                                                                     $rule.Style.Font.Color.color           = $ForeGroundColor     }
    if     ($PSBoundParameters.ContainsKey("BackgroundColor"  )   ) {if ($BackgroundColor -is [string])      {$BackgroundColor = [System.Drawing.Color]::$BackgroundColor }
                                                                     $rule.Style.Fill.BackgroundColor.color = $BackgroundColor     }
    if     ($PSBoundParameters.ContainsKey("BackgroundPattern")   ) {$rule.Style.Fill.PatternType           = $BackgroundPattern   }
    if     ($PSBoundParameters.ContainsKey("PatternColor"     )   ) {if ($PatternColor -is [string])         {$PatternColor = [System.Drawing.Color]::$PatternColor }
                                                                     $rule.Style.Fill.PatternColor.color    = $PatternColor        }
    #endregion
    #Allow further tweaking by returning the rule, if passthru specified
    if     ($Passthru)  {$rule}
}