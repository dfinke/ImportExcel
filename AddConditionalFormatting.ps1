Function Add-ConditionalFormatting {
    <#
      .Synopsis
        Adds contitional formatting to worksheet.
      .Example
        $excel = $avdata | Export-Excel -Path (Join-path $FilePath "\Machines.XLSX" ) -WorksheetName "Server Anti-Virus" -AutoSize -FreezeTopRow -AutoFilter -PassThru

        Add-ConditionalFormatting -WorkSheet $excel.Workbook.Worksheets[1] -Address "b2:b1048576" -ForeGroundColor "RED"     -RuleType ContainsText -ConditionValue "2003"
        Add-ConditionalFormatting -WorkSheet $excel.Workbook.Worksheets[1] -Address "i2:i1048576" -ForeGroundColor "RED"     -RuleType ContainsText -ConditionValue "Disabled"
        $excel.Workbook.Worksheets[1].Cells["D1:G1048576"].Style.Numberformat.Format = [cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern
        $excel.Workbook.Worksheets[1].Row(1).style.font.bold = $true
        $excel.Save() ; $excel.Dispose()

        Here Export-Excel is called with the -passThru parameter so the Excel Package object is stored in $Excel
        The desired worksheet is selected and the then columns B and i are conditially formatted (excluding the top row) to show red text if
        the columns contain "2003" or "Disabled respectively. A fixed date formats are then applied to columns D..G, and the top row is formatted.
        Finally the workbook is saved and the Excel object closed.
      .Example
        C:\> $r = Add-ConditionalFormatting -WorkSheet $excel.Workbook.Worksheets[1] -Range "B1:B100" -ThreeIconsSet Flags -Passthru
        C:\> $r.Reverse = $true ;   $r.Icon1.Type = "Num"; $r.Icon2.Type = "Num" ; $r.Icon2.value = 100 ; $r.Icon3.type = "Num" ;$r.Icon3.value = 1000

        Again Export excel has been called with -passthru leaving a package object in $Excel
        This time B1:B100 has been conditionally formatted with 3 icons, using the flags icon set.
        Add-ConditionalFormatting does not provide access to every option in the formatting rule, so passthru has been used and the
        rule is to apply the flags in reverse order, and boundaries for the number which will set the split are set to 100 and 1000
    #>
    Param (
        #The worksheet where the format is to be applied
        [Parameter(Mandatory = $true, ParameterSetName = "NamedRule")]
        [Parameter(Mandatory = $true, ParameterSetName = "DataBar")]
        [Parameter(Mandatory = $true, ParameterSetName = "ThreeIconSet")]
        [Parameter(Mandatory = $true, ParameterSetName = "FourIconSet")]
        [Parameter(Mandatory = $true, ParameterSetName = "FiveIconSet")]
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet ,
        #The area of the worksheet where the format is to be applied
        [Parameter(Mandatory = $true, ParameterSetName = "NamedRule")]
        [Parameter(Mandatory = $true, ParameterSetName = "DataBar")]
        [Parameter(Mandatory = $true, ParameterSetName = "ThreeIconSet")]
        [Parameter(Mandatory = $true, ParameterSetName = "FourIconSet")]
        [Parameter(Mandatory = $true, ParameterSetName = "FiveIconSet")]
        [OfficeOpenXml.ExcelAddress]$Range ,
        #One or more row(s), column(s) and/or block(s) of cells to format
        [Parameter(Mandatory = $true, ParameterSetName = "NamedRuleAddress")]
        [Parameter(Mandatory = $true, ParameterSetName = "DataBarAddress")]
        [Parameter(Mandatory = $true, ParameterSetName = "ThreeIconSetAddress")]
        [Parameter(Mandatory = $true, ParameterSetName = "FourIconSetAddress")]
        [Parameter(Mandatory = $true, ParameterSetName = "FiveIconSetAddress")]
        $Address ,
        #One of the standard named rules - Top / Bottom / Less than / Greater than / Contains etc.
        [Parameter(Mandatory = $true, ParameterSetName = "NamedRule", Position = 3)]
        [Parameter(Mandatory = $true, ParameterSetName = "NamedRuleAddress", Position = 3)]
        [OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType]$RuleType ,
        #Text colour for matching objects
        [Alias("ForeGroundColour")]
        [System.Drawing.Color]$ForeGroundColor,
        #colour for databar type charts
        [Parameter(Mandatory = $true, ParameterSetName = "DataBar")]
        [Parameter(Mandatory = $true, ParameterSetName = "DataBarAddress")]
        [Alias("DataBarColour")]
        [System.Drawing.Color]$DataBarColor,
        #One of the three-icon set types (e.g. Traffic Lights)
        [Parameter(Mandatory = $true, ParameterSetName = "ThreeIconSet")]
        [Parameter(Mandatory = $true, ParameterSetName = "ThreeIconSetAddress")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting3IconsSetType]$ThreeIconsSet,
        #A four-icon set name
        [Parameter(Mandatory = $true, ParameterSetName = "FourIconSet")]
        [Parameter(Mandatory = $true, ParameterSetName = "FourIconSetAddress")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting4IconsSetType]$FourIconsSet,
        #A five-icon set name
        [Parameter(Mandatory = $true, ParameterSetName = "FiveIconSet")]
        [Parameter(Mandatory = $true, ParameterSetName = "FiveIconSetAddress")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting5IconsSetType]$FiveIconsSet,
        #Use the icon set in reverse order
        [Parameter(ParameterSetName = "ThreeIconSet")]
        [Parameter(ParameterSetName = "ThreeIconSetAddress")]
        [Parameter(ParameterSetName = "FourIconSet")]
        [Parameter(ParameterSetName = "FourIconSetAddress")]
        [Parameter(ParameterSetName = "FiveIconSet")]
        [Parameter(ParameterSetName = "FiveIconSetAddress")]
        [switch]$Reverse,
        #A value for the condition (e.g. "2000" if the test is 'lessthan 2000')
        [string]$ConditionValue,
        #A second value for the conditions like "between x and Y"
        [string]$ConditionValue2,
        #Background colour for matching items
        [System.Drawing.Color]$BackgroundColor,
        #Background pattern for matching items
        [OfficeOpenXml.Style.ExcelFillStyle]$BackgroundPattern = [OfficeOpenXml.Style.ExcelFillStyle]::None ,
        #Secondary colour when a background pattern requires it
        [System.Drawing.Color]$PatternColor,
        #Sets the numeric format for matching items
        $NumberFormat,
        #Put matching items in bold face
        [switch]$Bold,
        #Put matching items in italic
        [switch]$Italic,
        #Underline matching items
        [switch]$Underline,
        #Strikethrough text of matching items
        [switch]$StrikeThru,
        #If specified pass the rule back to the caller to allow additional customization.
        [switch]$Passthru
    )
     
    #Allow conditional formatting to work like Set-Format (with single ADDRESS parameter), split it to get worksheet and range of cells.
    If ($Address -and -not $WorkSheet -and -not $Range) {
        $WorkSheet = $Address.Worksheet[0]
        $Range     = $Address.Address
    }
    #region Create a rule of the right type
    if     ($PSBoundParameters.ContainsKey("ThreeIconsSet" )      ) {$rule =  $WorkSheet.ConditionalFormatting.AddThreeIconSet($Range , $ThreeIconsSet)}
    elseif ($PSBoundParameters.ContainsKey("FourIconsSet"  )      ) {$rule =  $WorkSheet.ConditionalFormatting.AddFourIconSet( $Range , $FourIconsSet) }
    elseif ($PSBoundParameters.ContainsKey("FiveIconsSet"  )      ) {$rule =  $WorkSheet.ConditionalFormatting.AddFiveIconSet( $Range , $FiveIconsSet) }
    elseif ($PSBoundParameters.ContainsKey("DataBarColor"  )      ) {$rule =  $WorkSheet.ConditionalFormatting.AddDatabar(     $Range , $DataBarColor) }
    else                                                            {$rule = ($WorkSheet.ConditionalFormatting)."Add$RuleType"($Range)}
    if     ($PSBoundParameters.ContainsKey("Reverse"       )      ) {$rule.reverse = [boolean]$Reverse}
    #endregion
    #region set the rule conditions
    if     ($PSBoundParameters.ContainsKey("ConditionValue") -and 
            $RuleType -match "Top|Botom"                          ) {$rule.Rank     = $ConditionValue }
    if     ($PSBoundParameters.ContainsKey("ConditionValue") -and 
            $RuleType -match "StdDev"                             ) {$rule.StdDev   = $ConditionValue }
    if     ($PSBoundParameters.ContainsKey("ConditionValue") -and 
            $RuleType -match "Than|Equal|Expression"              ) {$rule.Formula  = $ConditionValue }
    if     ($PSBoundParameters.ContainsKey("ConditionValue") -and 
            $RuleType -match "Text|With"                          ) {$rule.Text     = $ConditionValue }
    if     ($PSBoundParameters.ContainsKey("ConditionValue") -and
            $PSBoundParameters.ContainsKey("ConditionValue") -and 
            $RuleType -match "Between"                            ) {
                                                                     $rule.Formula  = $ConditionValue; 
                                                                     $rule.Formula2 = $ConditionValue2
    }
    #endregion
    #region set the rule format 
    if     ($PSBoundParameters.ContainsKey("NumberFormat"  )      ) {$rule.Style.NumberFormat.Format        = (Expand-NumberFormat  $NumberFormat)             }
    if     ($Underline                                            ) {$rule.Style.Font.Underline             = [OfficeOpenXml.Style.ExcelUnderLineType]::Single }
    elseif ($PSBoundParameters.ContainsKey("Underline"     )      ) {$rule.Style.Font.Underline             = [OfficeOpenXml.Style.ExcelUnderLineType]::None   }
    if     ($PSBoundParameters.ContainsKey("Bold"          )      ) {$rule.Style.Font.Bold                  = [boolean]$Bold       }
    if     ($PSBoundParameters.ContainsKey("Italic"        )      ) {$rule.Style.Font.Italic                = [boolean]$Italic     }
    if     ($PSBoundParameters.ContainsKey("StrikeThru")          ) {$rule.Style.Font.Strike                = [boolean]$StrikeThru }
    if     ($PSBoundParameters.ContainsKey("ForeGroundColor"  )   ) {$rule.Style.Font.Color.color           = $ForeGroundColor     }
    if     ($PSBoundParameters.ContainsKey("BackgroundColor"  )   ) {$rule.Style.Fill.BackgroundColor.color = $BackgroundColor     }
    if     ($PSBoundParameters.ContainsKey("BackgroundPattern")   ) {$rule.Style.Fill.PatternType           = $BackgroundPattern   }
    if     ($PSBoundParameters.ContainsKey("PatternColor"     )   ) {$rule.Style.Fill.PatternColor.color    = $PatternColor        }
    #endregion 
    #Allow further tweaking by returning the rule, if passthru specified 
    if     ($Passthru)  {$rule}
}