Function Add-ConditionalFormatting {
<#
.Synopsis
    Adds contitional formatting to worksheet
.Example
    $excel = $avdata | Export-Excel -Path (Join-path $FilePath "\Machines.XLSX" ) -WorksheetName "Server Anti-Virus" -AutoSize -FreezeTopRow -AutoFilter -PassThru

    Add-ConditionalFormatting -WorkSheet $excel.Workbook.Worksheets[1] -Address "b":b1048576" -ForeGroundColor "RED"     -RuleType ContainsText -ConditionValue "2003"
    Add-ConditionalFormatting -WorkSheet $excel.Workbook.Worksheets[1] -Address "i2:i1048576" -ForeGroundColor "RED"     -RuleType ContainsText -ConditionValue "Disabled"
    $excel.Workbook.Worksheets[1].Cells["D1:G1048576"].Style.Numberformat.Format = [cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern
    $excel.Workbook.Worksheets[1].Row(1).style.font.bold = $true
    $excel.Save() ; $excel.Dispose()

    Here Export-Excel is called with the -passThru parameter so the Excel Package object is stored in $Excel
    The desired worksheet is selected and the then columns B and i are conditially formatted (excluding the top row) to show
    Fixed formats are then applied to dates in columns D..G and the top row is formatted
    Finally the workbook is saved and the Excel closed.

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
        #One or more row(s), Column(s) and/or block(s) of cells to format
        [Parameter(Mandatory = $true, ParameterSetName = "NamedRuleAddress")]
        [Parameter(Mandatory = $true, ParameterSetName = "DataBarAddress")]
        [Parameter(Mandatory = $true, ParameterSetName = "ThreeIconSetAddress")]
        [Parameter(Mandatory = $true, ParameterSetName = "FourIconSetAddress")]
        [Parameter(Mandatory = $true, ParameterSetName = "FiveIconSetAddress")]
        $Address ,
        #One of the standard named rules - Top / Bottom / Less than / Greater than / Contains etc
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
        #A value for the condition (e.g. "2000" if the test is 'lessthan 2000')
        [string]$ConditionValue,
        #A second value for the conditions like between x and Y
        [string]$ConditionValue2,
        #Background colour for matching items
        [System.Drawing.Color]$BackgroundColor,
        #Background pattern for matching items
        [OfficeOpenXml.Style.ExcelFillStyle]$BackgroundPattern = [OfficeOpenXml.Style.ExcelFillStyle]::Solid,
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
        [switch]$StrikeThru
    )
    #Allow add conditional formatting to work like Set-Format (with single ADDRESS parameter) split it to get worksheet and Range of cells.  
    If ($Address -and -not $WorkSheet -and -not $Range) {
        $WorkSheet = $Address.Worksheet[0]
        $Range     = $Address.Address 
    }    
    If ($ThreeIconsSet) {$rule = $WorkSheet.ConditionalFormatting.AddThreeIconSet($Range , $ThreeIconsSet)}
    elseif ($FourIconsSet) {$rule = $WorkSheet.ConditionalFormatting.AddFourIconSet( $Range , $FourIconsSet) }
    elseif ($FiveIconsSet) {$rule = $WorkSheet.ConditionalFormatting.AddFiveIconSet( $Range , $IconType)     }
    elseif ($DataBarColor) {$rule = $WorkSheet.ConditionalFormatting.AddDatabar(     $Range , $DataBarColor) }
    else {                  $rule = ($WorkSheet.ConditionalFormatting)."Add$RuleType"($Range)}

    if ($ConditionValue -and $RuleType -match "Top|Botom") {$rule.Rank = $ConditionValue }
    if ($ConditionValue -and $RuleType -match "StdDev") {$rule.StdDev = $ConditionValue }
    if ($ConditionValue -and $RuleType -match "Than|Equal|Expression") {$rule.Formula = $ConditionValue }
    if ($ConditionValue -and $RuleType -match "Text|With") {$rule.Text = $ConditionValue }
    if ($ConditionValue -and
        $ConditionValue2 -and $RuleType -match "Between") {
        $rule.Formula = $ConditionValue
        $rule.Formula2 = $ConditionValue2
    }

    if ($NumberFormat) {$rule.Style.NumberFormat.Format = $NumberFormat }
    if ($Underline) {$rule.Style.Font.Underline = [OfficeOpenXml.Style.ExcelUnderLineType]::Single }
    if ($Bold) {$rule.Style.Font.Bold = $true}
    if ($Italic) {$rule.Style.Font.Italic = $true}
    if ($StrikeThru) {$rule.Style.Font.Strike = $true}
    if ($ForeGroundColor) {$rule.Style.Font.Color.color = $ForeGroundColor   }
    if ($BackgroundColor) {$rule.Style.Fill.BackgroundColor.color = $BackgroundColor   }
    if ($BackgroundPattern) {$rule.Style.Fill.PatternType = $BackgroundPattern }
    if ($PatternColor) {$rule.Style.Fill.PatternColor.color = $PatternColor      }
}