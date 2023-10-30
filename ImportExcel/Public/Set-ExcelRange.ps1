function Set-ExcelRange {
    [CmdletBinding()]
    [Alias("Set-Format")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',Justification='Does not change system state')]
    param(
        [Parameter(ValueFromPipeline = $true,Position=0)]
        [Alias("Address")]
        $Range ,
        [OfficeOpenXml.ExcelWorksheet]$Worksheet ,
        [Alias("NFormat")]
        $NumberFormat,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderAround,
        $BorderColor=[System.Drawing.Color]::Black,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderBottom,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderTop,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderLeft,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderRight,
        [Alias('ForegroundColor')]
        $FontColor,
        $Value,
        $Formula,
        [Switch]$ArrayFormula,
        [Switch]$ResetFont,
        [Switch]$Bold,
        [Switch]$Italic,
        [Switch]$Underline,
        [OfficeOpenXml.Style.ExcelUnderLineType]$UnderLineType = [OfficeOpenXml.Style.ExcelUnderLineType]::Single,
        [Switch]$StrikeThru,
        [OfficeOpenXml.Style.ExcelVerticalAlignmentFont]$FontShift,
        [String]$FontName,
        [float]$FontSize,
        $BackgroundColor,
        [OfficeOpenXml.Style.ExcelFillStyle]$BackgroundPattern = [OfficeOpenXml.Style.ExcelFillStyle]::Solid ,
        [Alias("PatternColour")]
        $PatternColor,
        [Switch]$WrapText,
        [OfficeOpenXml.Style.ExcelHorizontalAlignment]$HorizontalAlignment,
        [OfficeOpenXml.Style.ExcelVerticalAlignment]$VerticalAlignment,
        [ValidateRange(-90, 90)]
        [int]$TextRotation ,
        [Alias("AutoFit")]
        [Switch]$AutoSize,
        [float]$Width,
        [float]$Height,
        [Alias('Hide')]
        [Switch]$Hidden,
        [Switch]$Locked,
        [Switch]$Merge
    )
    process {
        if  ($Range -is [Array])  {
            $null = $PSBoundParameters.Remove("Range")
            $Range | Set-ExcelRange @PSBoundParameters
        }
        else {
            #We should accept, a worksheet and a name of a range or a cell address; a table; the address of a table; a named range; a row, a column or .Cells[ ]
            if ($Range -is [OfficeOpenXml.Table.ExcelTable]) {$Range = $Range.Address}
            elseif ($Worksheet -and ($Range -is [string] -or $Range -is [OfficeOpenXml.ExcelAddress])) {
                $Range = $Worksheet.Cells[$Range]
            }
            elseif ($Range -is [string]) {Write-Warning -Message "The range parameter you have specified also needs a worksheet parameter." ;return}
            #else we assume $Range is a range.
            if ($ClearAll)  {
                $Range.Clear()
            }
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
            if ($PSBoundParameters.ContainsKey('Merge')) {
                $Range.Merge   = [boolean]$Merge
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
                        ForEach-Object {$Range.Worksheet.Row($_).Height = $Height }
                }
                else {Write-Warning -Message ("Can set the height of a row or a range but not a {0} object" -f ($Range.GetType().name)) }
            }
            if ($Autosize -and -not $env:NoAutoSize) {
                try {
                    if ($Range -is [OfficeOpenXml.ExcelColumn]) {$Range.AutoFit() }
                    elseif ($Range -is [OfficeOpenXml.ExcelRange] ) {
                        $Range.AutoFitColumns()

                    }
                    else {Write-Warning -Message ("Can autofit a column or a range but not a {0} object" -f ($Range.GetType().name)) }
                }
                catch {Write-Warning -Message "Failed autosizing columns of worksheet '$WorksheetName': $_"}
            }
            elseif ($AutoSize) {Write-Warning -Message "Auto-fitting columns is not available with this OS configuration." }
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
