Function Set-ExcelRange {
    <#
      .SYNOPSIS
        Applies number, font, alignment and/or color formatting, values or formulas to a range of Excel cells.
      .DESCRIPTION
        Set-ExcelRange was created to set the style elements for a range of cells,
        this includes auto-sizing and hiding, setting font elements (Name, Size,
        Bold, Italic, Underline & UnderlineStyle and Subscript & SuperScript),
        font and background colors, borders, text wrapping, rotation, alignment
        within cells, and number format.
        It was orignally named "Set-Format", but it has been extended to set
        Values, Formulas and ArrayFormulas (sometimes called Ctrl-shift-Enter
        [CSE] formulas); because of this, the name has become Set-ExcelRange
        but the old name of Set-Format is preserved as an alias.
      .EXAMPLE
        $sheet.Column(3) | Set-ExcelRange -HorizontalAlignment Right -NumberFormat "#,###" -AutoFit

        Selects column 3 from a sheet object (within a workbook object, which
        is a child of the ExcelPackage object) and passes it to Set-ExcelRange
        which formats numbers as a integers with comma-separated groups,
        aligns it right, and auto-fits the column to the contents.
      .EXAMPLE
        Set-ExcelRange -Range $sheet.Cells["E1:H1048576"]  -HorizontalAlignment Right -NumberFormat "#,###"

        Instead of piping the address, this version specifies a block of cells
        and applies similar formatting.
      .EXAMPLE
        Set-ExcelRange $excel.Workbook.Worksheets[1].Tables["Processes"] -Italic

        This time instead of specifying a range of cells, a table is selected
        by name and formatted as italic.
    #>
    [cmdletbinding()]
    [Alias("Set-Format")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',Justification='Does not change system state')]
    Param   (
        #One or more row(s), Column(s) and/or block(s) of cells to format.
        [Parameter(ValueFromPipeline = $true,Position=0)]
        [Alias("Address")]
        $Range ,
        #The worksheet where the format is to be applied.
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet ,
        #Number format to apply to cells e.g. "dd/MM/yyyy HH:mm", "£#,##0.00;[Red]-£#,##0.00", "0.00%" , "##/##" , "0.0E+0" etc.
        [Alias("NFormat")]
        $NumberFormat,
        #Style of border to draw around the range.
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderAround,
        #Color of the border.
        $BorderColor=[System.Drawing.Color]::Black,
        #Style for the bottom border.
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderBottom,
        #Style for the top border.
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderTop,
        #Style for the left border.
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderLeft,
        #Style for the right border.
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderRight,
        #Colour for the text - if none is specified it will be left as it is.
        $FontColor,
        #Value for the cell.
        $Value,
        #Formula for the cell.
        $Formula,
        #Specifies formula should be an array formula (a.k.a CSE [ctrl-shift-enter] formula).
        [Switch]$ArrayFormula,
        #Clear Bold, Italic, StrikeThrough and Underline and set color to Black.
        [Switch]$ResetFont,
        #Make text bold; use -Bold:$false to remove bold.
        [Switch]$Bold,
        #Make text italic;  use -Italic:$false to remove italic.
        [Switch]$Italic,
        #Underline the text using the underline style in -UnderlineType;  use -Underline:$false to remove underlining.
        [Switch]$Underline,
         #Specifies whether underlining should be single or double, normal or accounting mode. The default is "Single".
        [OfficeOpenXml.Style.ExcelUnderLineType]$UnderLineType = [OfficeOpenXml.Style.ExcelUnderLineType]::Single,
        #Strike through text; use -Strikethru:$false to remove Strike through
        [Switch]$StrikeThru,
        #Subscript or Superscript (or none).
        [OfficeOpenXml.Style.ExcelVerticalAlignmentFont]$FontShift,
        #Font to use - Excel defaults to Calibri.
        [String]$FontName,
        #Point size for the text.
        [float]$FontSize,
        #Change background color.
        $BackgroundColor,
        #Background pattern - Solid by default.
        [OfficeOpenXml.Style.ExcelFillStyle]$BackgroundPattern = [OfficeOpenXml.Style.ExcelFillStyle]::Solid ,
        #Secondary color for background pattern.
        [Alias("PatternColour")]
        $PatternColor,
        #Turn on Text-Wrapping; use -WrapText:$false to turn off wrapping.
        [Switch]$WrapText,
        #Position cell contents to Left, Right, Center etc. default is 'General'.
        [OfficeOpenXml.Style.ExcelHorizontalAlignment]$HorizontalAlignment,
        #Position cell contents to Top, Bottom or Center.
        [OfficeOpenXml.Style.ExcelVerticalAlignment]$VerticalAlignment,
        #Degrees to rotate text. Up to +90 for anti-clockwise ("upwards"), or to -90 for clockwise.
        [ValidateRange(-90, 90)]
        [int]$TextRotation ,
        #Autofit cells to width  (columns or ranges only).
        [Alias("AutoFit")]
        [Switch]$AutoSize,
        #Set cells to a fixed width (columns or ranges only), ignored if Autosize is specified.
        [float]$Width,
        #Set cells to a fixed height  (rows or ranges only).
        [float]$Height,
        #Hide a row or column  (not a range); use -Hidden:$false to unhide.
        [Switch]$Hidden,
        #Locks cells. Cells are locked by default use -locked:$false on the whole sheet and then lock specific ones, and enable protection on the sheet.
        [Switch]$Locked
    )
    process {
        if  ($Range -is [Array])  {
            $null = $PSBoundParameters.Remove("Range")
            $Range | Set-ExcelRange @PSBoundParameters
        }
        else {
            #We should accept, a worksheet and a name of a range or a cell address; a table; the address of a table; a named range; a row, a column or .Cells[ ]
            if ($Range -is [OfficeOpenXml.Table.ExcelTable]) {$Range = $Range.Address}
            elseif ($WorkSheet -and ($Range -is [string] -or $Range -is [OfficeOpenXml.ExcelAddress])) {
                $Range = $WorkSheet.Cells[$Range]
            }
            elseif ($Range -is [string]) {Write-Warning -Message "The range pararameter you have specified also needs a worksheet parameter." ;return}
            #else we assume Range is a range.
            if ($ResetFont) {
                $Range.Style.Font.Color.SetColor( ([System.Drawing.Color]::Black))
                $Range.Style.Font.Bold          = $false
                $Range.Style.Font.Italic        = $false
                $Range.Style.Font.UnderLine     = $false
                $Range.Style.Font.Strike        = $false
                $Range.Style.Font.VerticalAlign = [OfficeOpenXml.Style.ExcelVerticalAlignmentFont]::None
            }
            if ($PSBoundParameters.ContainsKey('Underline')) {
                $Range.Style.Font.UnderLine      = [boolean]$Underline
                $Range.Style.Font.UnderLineType  = $UnderLineType
            }
            if ($PSBoundParameters.ContainsKey('Bold')) {
                $Range.Style.Font.Bold           = [boolean]$bold
            }
            if ($PSBoundParameters.ContainsKey('Italic')) {
                $Range.Style.Font.Italic         = [boolean]$italic
            }
            if ($PSBoundParameters.ContainsKey('StrikeThru')) {
                $Range.Style.Font.Strike         = [boolean]$StrikeThru
            }
            if ($PSBoundParameters.ContainsKey('FontSize')){
                $Range.Style.Font.Size           = $FontSize
            }
            if ($PSBoundParameters.ContainsKey('FontName')){
                $Range.Style.Font.Name           = $FontName
            }
            if ($PSBoundParameters.ContainsKey('FontShift')){
                $Range.Style.Font.VerticalAlign  = $FontShift
            }
            if ($PSBoundParameters.ContainsKey('FontColor')){
                if ($FontColor -is [string]) {$FontColor = [System.Drawing.Color]::$FontColor }
                $Range.Style.Font.Color.SetColor(  $FontColor)
            }
            if ($PSBoundParameters.ContainsKey('TextRotation')) {
                $Range.Style.TextRotation        = $TextRotation
            }
            if ($PSBoundParameters.ContainsKey('WrapText')) {
                $Range.Style.WrapText            = [boolean]$WrapText
            }
            if ($PSBoundParameters.ContainsKey('HorizontalAlignment')) {
                $Range.Style.HorizontalAlignment = $HorizontalAlignment
            }
            if ($PSBoundParameters.ContainsKey('VerticalAlignment')) {
                $Range.Style.VerticalAlignment   = $VerticalAlignment
            }
            if ($PSBoundParameters.ContainsKey('Value')) {
                if ($Value -match '^=')      {$PSBoundParameters["Formula"] = $Value -replace '^=','' }
                else {
                    $Range.Value = $Value
                    if ($Value -is [datetime])  { $Range.Style.Numberformat.Format = 'm/d/yy h:mm' }# This is not a custom format, but a preset recognized as date and localized. It might be overwritten in a moment
                    if ($Value -is [timespan])  { $Range.Style.Numberformat.Format = '[h]:mm:ss'   }
                }
            }
            if ($PSBoundParameters.ContainsKey('Formula')) {
                if ($ArrayFormula) {$Range.CreateArrayFormula(($Formula -replace '^=','')) }
                else               {$Range.Formula         =  ($Formula -replace '^=','')  }
            }
            if ($PSBoundParameters.ContainsKey('NumberFormat')) {
                $Range.Style.Numberformat.Format = (Expand-NumberFormat $NumberFormat)
            }
            if ($BorderColor -is [string]) {$BorderColor = [System.Drawing.Color]::$BorderColor }
            if ($PSBoundParameters.ContainsKey('BorderAround')) {
                $Range.Style.Border.BorderAround($BorderAround, $BorderColor)
            }
            if ($PSBoundParameters.ContainsKey('BorderBottom')) {
                $Range.Style.Border.Bottom.Style=$BorderBottom
                $Range.Style.Border.Bottom.Color.SetColor($BorderColor)
            }
            if ($PSBoundParameters.ContainsKey('BorderTop')) {
                $Range.Style.Border.Top.Style=$BorderTop
                $Range.Style.Border.Top.Color.SetColor($BorderColor)
            }
            if ($PSBoundParameters.ContainsKey('BorderLeft')) {
                $Range.Style.Border.Left.Style=$BorderLeft
                $Range.Style.Border.Left.Color.SetColor($BorderColor)
            }
            if ($PSBoundParameters.ContainsKey('BorderRight')) {
                $Range.Style.Border.Right.Style=$BorderRight
                $Range.Style.Border.Right.Color.SetColor($BorderColor)
            }
            if ($PSBoundParameters.ContainsKey('BackgroundColor')) {
                $Range.Style.Fill.PatternType = $BackgroundPattern
                if ($BackgroundColor -is [string]) {$BackgroundColor = [System.Drawing.Color]::$BackgroundColor }
                $Range.Style.Fill.BackgroundColor.SetColor($BackgroundColor)
                if ($PatternColor) {
                    if ($PatternColor -is [string]) {$PatternColor = [System.Drawing.Color]::$PatternColor }
                    $Range.Style.Fill.PatternColor.SetColor( $PatternColor)
                }
            }
            if ($PSBoundParameters.ContainsKey('Height')) {
                if ($Range -is [OfficeOpenXml.ExcelRow]   ) {$Range.Height = $Height }
                elseif ($Range -is [OfficeOpenXml.ExcelRange] ) {
                    ($Range.Start.Row)..($Range.Start.Row + $Range.Rows) |
                        ForEach-Object {$Range.WorkSheet.Row($_).Height = $Height }
                }
                else {Write-Warning -Message ("Can set the height of a row or a range but not a {0} object" -f ($Range.GetType().name)) }
            }
            if ($Autosize) {
                if ($Range -is [OfficeOpenXml.ExcelColumn]) {$Range.AutoFit() }
                elseif ($Range -is [OfficeOpenXml.ExcelRange] ) {
                    $Range.AutoFitColumns()
                }
                else {Write-Warning -Message ("Can autofit a column or a range but not a {0} object" -f ($Range.GetType().name)) }

            }
            elseif ($PSBoundParameters.ContainsKey('Width')) {
                if ($Range -is [OfficeOpenXml.ExcelColumn]) {$Range.Width = $Width}
                elseif ($Range -is [OfficeOpenXml.ExcelRange] ) {
                    ($Range.Start.Column)..($Range.Start.Column + $Range.Columns - 1) |
                        ForEach-Object {
                            #$ws.Column($_).Width = $Width
                            $Range.Worksheet.Column($_).Width = $Width
                        }
                }
                else {Write-Warning -Message ("Can set the width of a column or a range but not a {0} object" -f ($Range.GetType().name)) }
            }
            if ($PSBoundParameters.ContainsKey('Hidden')) {
                if ($Range -is [OfficeOpenXml.ExcelRow] -or
                    $Range -is [OfficeOpenXml.ExcelColumn]  ) {$Range.Hidden = [boolean]$Hidden}
                else {Write-Warning -Message ("Can hide a row or a column but not a {0} object" -f ($Range.GetType().name)) }
            }
            if ($PSBoundParameters.ContainsKey('Locked')) {
                $Range.Style.Locked=$Locked
            }
        }
    }
}

Function NumberFormatCompletion {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    $numformats = [ordered]@{
        "General"       = "General"      # format ID  0
        "Number"        = "0.00"         # format ID  2
        "Percentage"    = "0.00%"        # format ID 10
        "Scientific"    = "0.00E+00"     # format ID 11
        "Fraction"      = "# ?/?"        # format ID 12
        "Short Date"    = "Localized"    # format ID 14 - will be translated to "mm-dd-yy"     which is localized on load by Excel.
        "Short Time"    = "Localized"    # format ID 20 - will be translated to "h:mm"         which is localized on load by Excel.
        "Long Time"     = "Localized"    # format ID 21 - will be translated to "h:mm:ss"      which is localized on load by Excel.
        "Date-Time"     = "Localized"    # format ID 22 - will be translated to "m/d/yy h:mm"  which is localized on load by Excel.
        "Currency"      = [cultureinfo]::CurrentCulture.NumberFormat.CurrencySymbol + "#,##0.00"
        "Text"          = "@"              # format ID 49
        "h:mm AM/PM"    = "h:mm AM/PM"     # format ID 18
        "h:mm:ss AM/PM" = "h:mm:ss AM/PM"  # format ID 19
        "mm:ss"         = "mm:ss"          # format ID 45
        "[h]:mm:ss"     = "Elapsed hours"  # format ID 46
        "mm:ss.0"       = "mm:ss.0"        # format ID 47
        "d-mmm-yy"      = "Localized"      # format ID 15 which is localized on load by Excel.
        "d-mmm"         = "Localized"      # format ID 16 which is localized on load by Excel.
        "mmm-yy"        = "mmm-yy"         # format ID 17 which is localized on load by Excel.
        "0"             = "Whole number"                       # format ID  1
        "0.00"          = "Number, 2 decimals"                 # format ID  2 or "number"
        "#,##0"         = "Thousand separators"                # format ID  3
        "#,##0.00"      = "Thousand separators and 2 decimals" # format ID  4
        "#,"            = "Whole thousands"
        "#.0,,"         = "Millions, 1 Decimal"
        "0%"            = "Nearest whole percentage"           # format ID  9
        "0.00%"         = "Percentage with decimals"           # format ID 10 or "Percentage"
        "00E+00"        = "Scientific"                         # format ID 11 or "Scientific"
        "# ?/?"         = "One Digit fraction"                 # format ID 12 or "Fraction"
        "# ??/??"       = "Two Digit fraction"                 # format ID 13
        "@"             = "Text"                               # format ID 49 or "Text"
    }
    $numformats.keys.where({$_ -like "$wordToComplete*"} ) | ForEach-Object {
        New-Object -TypeName System.Management.Automation.CompletionResult -ArgumentList "'$_'" , $_ ,
        ([System.Management.Automation.CompletionResultType]::ParameterValue) , $numformats[$_]
    }
}
if (Get-Command -ErrorAction SilentlyContinue -name Register-ArgumentCompleter) {
    Register-ArgumentCompleter -CommandName Add-ConditionalFormatting  -ParameterName NumberFormat        -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Export-Excel               -ParameterName NumberFormat        -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelRange             -ParameterName NumberFormat        -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelColumn            -ParameterName NumberFormat        -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelRow               -ParameterName NumberFormat        -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Add-PivotTable             -ParameterName PivotNumberFormat   -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName New-PivotTableDefinition   -ParameterName PivotNumberFormat   -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName New-ExcelChartDefinition   -ParameterName XAxisNumberformat   -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName New-ExcelChartDefinition   -ParameterName YAxisNumberformat   -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Add-ExcelChart             -ParameterName XAxisNumberformat   -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Add-ExcelChart             -ParameterName YAxisNumberformat   -ScriptBlock $Function:NumberFormatCompletion
}

Function Expand-NumberFormat {
    <#
      .SYNOPSIS
        Converts short names for number formats to the formatting strings used in Excel
      .DESCRIPTION
        Where you can type a number format you can write, for example, 'Short-Date'
        and the module will translate it into the format string used by Excel.
        Some formats, like Short-Date change how they are presented when Excel
        loads (so date will use the local ordering of year, month and Day). Other
        formats change how they appear when loaded with different cultures
        (depending on the country "," or "." or " " may be the thousand seperator
        although Excel always stores it as ",")
      .EXAMPLE
        Expand-NumberFormat percentage

        Returns "0.00%"
      .EXAMPLE
        Expand-NumberFormat Currency

        Returns the currency format specified in the local regional settings. This
        may not be the same as Excel uses.  The regional settings set the currency
        symbol and then whether it is before or after the number and separated with
        a space or not; for negative numbers the number may be wrapped in parentheses
        or a - sign might appear before or after the number and symbol.
        So this returns $#,##0.00;($#,##0.00) for English US, #,##0.00 €;€#,##0.00-
        for French. (Note some Eurozone countries write €1,23 and others 1,23€ )
        In French the decimal point will be rendered as a "," and the thousand
        separator as a space.
    #>
    [cmdletbinding()]
    [OutputType([String])]
    param  (
        #the format string to Expand
        $NumberFormat
    )
    switch ($NumberFormat) {
        "Currency"      {
            #https://msdn.microsoft.com/en-us/library/system.globalization.numberformatinfo.currencynegativepattern(v=vs.110).aspx
            $sign = [cultureinfo]::CurrentCulture.NumberFormat.CurrencySymbol
            switch ([cultureinfo]::CurrentCulture.NumberFormat.CurrencyPositivePattern) {
                0  {$pos = "$Sign#,##0.00"  ; break }
                1  {$pos = "#,##0.00$Sign"  ; break }
                2  {$pos = "$Sign #,##0.00" ; break }
                3  {$pos = "#,##0.00 $Sign" ; break }
            }
            switch ([cultureinfo]::CurrentCulture.NumberFormat.CurrencyPositivePattern) {
                0  {return "$pos;($Sign#,##0.00)"  }
                1  {return "$pos;-$Sign#,##0.00"   }
                2  {return "$pos;$Sign-#,##0.00"   }
                3  {return "$pos;$Sign#,##0.00-"   }
                4  {return "$pos;(#,##0.00$Sign)"  }
                5  {return "$pos;-#,##0.00$Sign"   }
                6  {return "$pos;#,##0.00-$Sign"   }
                7  {return "$pos;#,##0.00$Sign-"   }
                8  {return "$pos;-#,##0.00 $Sign"  }
                9  {return "$pos;-$Sign #,##0.00"  }
               10  {return "$pos;#,##0.00 $Sign-"  }
               11  {return "$pos;$Sign #,##0.00-"  }
               12  {return "$pos;$Sign -#,##0.00"  }
               13  {return "$pos;#,##0.00- $Sign"  }
               14  {return "$pos;($Sign #,##0.00)" }
               15  {return "$pos;(#,##0.00 $Sign)" }
            }
        }
        "Number"        {return  "0.00"       } # format id  2
        "Percentage"    {return  "0.00%"      } # format id 10
        "Scientific"    {return  "0.00E+00"   } # format id 11
        "Fraction"      {return  "# ?/?"      } # format id 12
        "Short Date"    {return  "mm-dd-yy"   } # format id 14 localized on load by Excel.
        "Short Time"    {return  "h:mm"       } # format id 20 localized on load by Excel.
        "Long Time"     {return  "h:mm:ss"    } # format id 21 localized on load by Excel.
        "Date-Time"     {return  "m/d/yy h:mm"} # format id 22 localized on load by Excel.
        "Text"          {return  "@"          } # format ID 49
        Default         {return  $NumberFormat}
    }
}
