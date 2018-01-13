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
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet ,
        #The area of the worksheet where the format is to be applied
        [OfficeOpenXml.ExcelAddress]$Range ,
        #One of the standard named rules - Top / Bottom / Less than / Greater than / Contains etc
        [Parameter(Mandatory=$true,ParameterSetName="NamedRule",Position=3)]
        [OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType]$RuleType ,
        #Text colour for matching objects
        [Alias("ForeGroundColour")]
        [System.Drawing.Color]$ForeGroundColor,
        #colour for databar type charts
        [Parameter(Mandatory=$true,ParameterSetName="DataBar")]
        [Alias("DataBarColour")]
        [System.Drawing.Color]$DataBarColor,
        #One of the three-icon set types (e.g. Traffic Lights)
        [Parameter(Mandatory=$true,ParameterSetName="ThreeIconSet")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting3IconsSetType]$ThreeIconsSet,
        #A four-icon set name
        [Parameter(Mandatory=$true,ParameterSetName="FourIconSet")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting4IconsSetType]$FourIconsSet,
        #A five-icon set name
        [Parameter(Mandatory=$true,ParameterSetName="FiveIconSet")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting5IconsSetType]$FiveIconsSet,
        #A value for the condition (e.g. "2000" if the test is 'lessthan 2000')
        [string]$ConditionValue,
        #A second value for the conditions like between x and Y
        [string]$ConditionValue2,
        #Background colour for matching items
        [System.Drawing.Color]$BackgroundColor,
        #Background pattern for matching items
        [OfficeOpenXml.Style.ExcelFillStyle]$BackgroundPattern =  [OfficeOpenXml.Style.ExcelFillStyle]::Solid,
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

    If      ($ThreeIconsSet)  {$rule =  $WorkSheet.ConditionalFormatting.AddThreeIconSet($Range , $ThreeIconsSet)}
    elseif  ($FourIconsSet)   {$rule =  $WorkSheet.ConditionalFormatting.AddFourIconSet( $Range , $FourIconsSet) }
    elseif  ($FiveIconsSet)   {$rule =  $WorkSheet.ConditionalFormatting.AddFiveIconSet( $Range , $IconType)     }
    elseif  ($DataBarColor)   {$rule =  $WorkSheet.ConditionalFormatting.AddDatabar(     $Range , $DataBarColor) }
    else    {                  $rule = ($WorkSheet.ConditionalFormatting)."Add$RuleType"($Range)}

    if ($ConditionValue   -and $RuleType -match "Top|Botom")             {$rule.Rank      = $ConditionValue }
    if ($ConditionValue   -and $RuleType -match "StdDev")                {$rule.StdDev    = $ConditionValue }
    if ($ConditionValue   -and $RuleType -match "Than|Equal|Expression") {$rule.Formula   = $ConditionValue }
    if ($ConditionValue   -and $RuleType -match "Text|With")             {$rule.Text      = $ConditionValue }
    if ($ConditionValue   -and
         $ConditionValue2 -and $RuleType -match "Between")               {$rule.Formula   = $ConditionValue
                                                                          $rule.Formula2  = $ConditionValue2}

    if ($NumberFormat)        {$rule.Style.NumberFormat.Format        = $NumberFormat }
    if ($Underline)           {$rule.Style.Font.Underline             = [OfficeOpenXml.Style.ExcelUnderLineType]::Single }
    if ($Bold)                {$rule.Style.Font.Bold                  = $true}
    if ($Italic)              {$rule.Style.Font.Italic                = $true}
    if ($StrikeThru)          {$rule.Style.Font.Strike                = $true}
    if ($ForeGroundColor)     {$rule.Style.Font.Color.color           = $ForeGroundColor   }
    if ($BackgroundColor)     {$rule.Style.Fill.BackgroundColor.color = $BackgroundColor   }
    if ($BackgroundPattern)   {$rule.Style.Fill.PatternType           = $BackgroundPattern }
    if ($PatternColor)        {$rule.Style.Fill.PatternColor.color    = $PatternColor      }
}

Function Set-Format {
<#
.SYNOPSIS
    Applies Number, font, alignment and colour formatting to a range of Excel Cells
.EXAMPLE
    $sheet.Column(3) | Set-Format -HorizontalAlignment Right -NumberFormat "#,###"
    Selects column 3 from a sheet object (within a workbook object, which is a child of the ExcelPackage object) and passes it to Set-Format which formats as an integer with comma seperated groups
.EXAMPLE
    Set-Format -Address $sheet.Cells["E1:H1048576"]  -HorizontalAlignment Right -NumberFormat "#,###"
    Instead of piping the address in this version specifies a block of cells and applies similar formatting

#>
    Param   (
        #One or more row(s), Column(s) and/or block(s) of cells to format
        [Parameter(ValueFromPipeline=$true)]
        [object[]]$Address ,
        #Number format to apply to cells e.g. "dd/MM/yyyy HH:mm", "£#,##0.00;[Red]-£#,##0.00", "0.00%" , "##/##" , "0.0E+0" etc
        [Alias("NFormat")]
        $NumberFormat,
        #Style of border to draw around the range
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderAround,
        #Colour for the text - if none specified it will be left as it it is
        [System.Drawing.Color]$FontColor,
        #Clear Bold, Italic, StrikeThrough and Underline and set colour to black
        [switch]$ResetFont,
        #Make text bold
        [switch]$Bold,
        #Make text italic
        [switch]$Italic,
        #Underline the text using the underline style in -underline type
        [switch]$Underline,
        #Should Underline use single or double, normal or accounting mode : default is single normal
        [OfficeOpenXml.Style.ExcelUnderLineType]$UnderLineType = [OfficeOpenXml.Style.ExcelUnderLineType]::Single,
        #StrikeThrough text
        [switch]$StrikeThru,
        #Subscript or superscript
        [OfficeOpenXml.Style.ExcelVerticalAlignmentFont]$FontShift,
        #Font to use - Excel defaults to Calibri
        [String]$FontName,
        #Point size for the text
        [float]$FontSize,
        #Change background colour
        [System.Drawing.Color]$BackgroundColor,
        #Background pattern - solid by default
        [OfficeOpenXml.Style.ExcelFillStyle]$BackgroundPattern =[OfficeOpenXml.Style.ExcelFillStyle]::Solid ,
        #Secondary colour for background pattern
        [Alias("PatternColour")]
        [System.Drawing.Color]$PatternColor,
        #Turn on text wrapping
        [switch]$WrapText,
        #Position cell contents to left, right or centre ...
        [OfficeOpenXml.Style.ExcelHorizontalAlignment]$HorizontalAlignment,
        #Position cell contents to top bottom or centre
        [OfficeOpenXml.Style.ExcelVerticalAlignment]$VerticalAlignment,
        #Degrees to rotate text. Up to +90 for anti-clockwise ("upwards"), or to -90 for clockwise.
        [ValidateRange(-90,90)]
        [int]$TextRotation ,
        #Autofit cells to width  (columns or ranges only)
        [switch]$AutoFit,
        #Set cells to a fixed width (columns or ranges only), ignored if Autofit is specified
        [float]$Width,
        #Set cells to a fixed hieght  (rows or ranges only)
        [float]$Height,
        #Hide a row or column  (not a range)
        [switch]$Hidden
    )
    process {
     Foreach ($range in $Address) {
        if ($ResetFont)           {$Range.Style.Font.Color.SetColor("Black")
                                   $Range.Style.Font.Bold          = $false
                                   $Range.Style.Font.Italic        = $false
                                   $Range.Style.Font.UnderLine     = $false
                                   $Range.Style.Font.Strike        = $false
        }
        if ($Underline)           {$Range.Style.Font.UnderLine     = $true
                                   $Range.Style.Font.UnderLineType  =$UnderLineType
        }
        if ($Bold)                {$Range.Style.Font.Bold          = $true                }
        if ($Italic)              {$Range.Style.Font.Italic        = $true                }
        if ($StrikeThru)          {$Range.Style.Font.Strike        = $true                }
        if ($FontShift)           {$Range.Style.Font.VerticalAlign = $FontShift           }
        if ($FontColor)           {$Range.Style.Font.Color.SetColor( $FontColor    )      }
        if ($BorderAround)         {$Range.Style.Border.BorderAround( $BorderAround )      }
        if ($NumberFormat)        {$Range.Style.Numberformat.Format= $NumberFormat        }
        if ($TextRotation)        {$Range.Style.TextRotation       = $TextRotation        }
        if ($WrapText)            {$Range.Style.WrapText           = $true                }
        if ($HorizontalAlignment) {$Range.Style.HorizontalAlignment= $HorizontalAlignment }
        if ($VerticalAlignment)   {$Range.Style.VerticalAlignment  = $VerticalAlignment   }

        if ($BackgroundColor)     {
                    $Range.Style.Fill.PatternType = $BackgroundPattern
                    $Range.Style.Fill.BackgroundColor.SetColor($BackgroundColor)
                    if ($PatternColor) {
                            $range.Style.Fill.PatternColor.SetColor( $PatternColor)
                    }
        }

        if     ($Height)  {
            if     ($Range -is [OfficeOpenXml.ExcelRow]   ) {$Range.Height = $Height }
            elseif ($Range -is [OfficeOpenXml.ExcelRange] ) {
                   ($range.Start.Row)..($range.Start.Row + $range.Rows) |
                                        ForEach-Object {$ws.Row($_).Height = $Height }
            }
            else   {Write-Warning -Message ("Can set the height of a row or a range but not a {0} object" -f ($Range.GetType().name)) }
        }
        if     ($AutoFit) {
            if     ($Range -is [OfficeOpenXml.ExcelColumn]) {$Range.AutoFit() }
            elseif ($Range -is [OfficeOpenXml.ExcelRange] ) {$Range.AutoFitColumns() }
            else   {Write-Warning -Message ("Can autofit a column or a range but not a {0} object" -f ($Range.GetType().name)) }

        }
        elseif ($Width)   {
            if     ($Range -is [OfficeOpenXml.ExcelColumn]) {$Range.Width = $Width}
            elseif ($Range -is [OfficeOpenXml.ExcelRange] ) {
                   ($range.Start.Column)..($range.Start.Column+ $range.Columns) |
                                      ForEach-Object {$ws.Column($_).Width = $Width}
            }
            else   {Write-Warning -Message ("Can set the width of a column or a range but not a {0} object" -f ($Range.GetType().name)) }
        }
        if     ($Hidden)  {
            if ($Range -is [OfficeOpenXml.ExcelRow] -or
                $Range -is [OfficeOpenXml.ExcelColumn]  ) {$Range.Hidden = $True}
            else {Write-Warning -Message ("Can hide a row or a column but not a {0} object" -f ($Range.GetType().name)) }
        }
      }
    }
}

#Argument completer for colours. If we have PS 5 or Tab expansion++ then we'll register it. Otherwise it does nothing.
Function ColorCompletion{
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    [System.Drawing.KnownColor].GetFields() | Where-Object {$_.IsStatic -and $_.name -like "$wordToComplete*" } |
         Sort-Object name | ForEach-Object {New-CompletionResult $_.name $_.name
    }
}

if (Get-Command -Name register-argumentCompleter -ErrorAction SilentlyContinue) {
    Register-ArgumentCompleter -CommandName Export-Excel               -ParameterName TitleBackgroundColor -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Add-ConditionalFormatting  -ParameterName ForeGroundColor      -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Add-ConditionalFormatting  -ParameterName DataBarColor         -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Add-ConditionalFormatting  -ParameterName BackgroundColor      -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Format                 -ParameterName FontColor            -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Format                 -ParameterName BackgroundColor      -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Format                 -ParameterName PatternColor         -ScriptBlock $Function:ColorCompletion
}