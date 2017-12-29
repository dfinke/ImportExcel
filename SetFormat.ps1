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
.EXAMPLE
    Set-Format -Indent 4 $sheet.Cells[C1:C10]
    Sets selected cells indent to 4. Valid range is 0-15
#>
    Param   (
        #One or more row(s), Column(s) and/or block(s) of cells to format
        [Parameter(ValueFromPipeline = $true)]
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
        [OfficeOpenXml.Style.ExcelFillStyle]$BackgroundPattern = [OfficeOpenXml.Style.ExcelFillStyle]::Solid ,
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
        [ValidateRange(-90, 90)]
        [int]$TextRotation ,
        #Autofit cells to width  (columns or ranges only)
        [switch]$AutoFit,
        #Set cells to a fixed width (columns or ranges only), ignored if Autofit is specified
        [float]$Width,
        #Set cells to a fixed hieght  (rows or ranges only)
        [float]$Height,
        #Hide a row or column  (not a range)
        [switch]$Hidden,
        #Indent cells in range
        [ValidateRange(0, 15)]
        [int]$Indent
    )
    process {
        Foreach ($range in $Address) {
            if ($ResetFont) {
                $Range.Style.Font.Color.SetColor("Black")
                $Range.Style.Font.Bold = $false
                $Range.Style.Font.Italic = $false
                $Range.Style.Font.UnderLine = $false
                $Range.Style.Font.Strike = $falsee
            }
            if ($Underline) {
                $Range.Style.Font.UnderLine = $true
                $Range.Style.Font.UnderLineType = $UnderLineType
            }
            if ($Bold) {$Range.Style.Font.Bold = $true                }
            if ($Italic) {$Range.Style.Font.Italic = $true                }
            if ($StrikeThru) {$Range.Style.Font.Strike = $true                }
            if ($FontShift) {$Range.Style.Font.VerticalAlign = $FontShift           }
            if ($FontColor) {$Range.Style.Font.Color.SetColor( $FontColor    )      }
            if ($BorderRound) {$Range.Style.Border.BorderAround( $BorderAround )      }
            if ($NumberFormat) {$Range.Style.Numberformat.Format = $NumberFormat        }
            if ($TextRotation) {$Range.Style.TextRotation = $TextRotation        }
            if ($WrapText) {$Range.Style.WrapText = $true                }
            if ($HorizontalAlignment) {$Range.Style.HorizontalAlignment = $HorizontalAlignment }
            if ($VerticalAlignment) {$Range.Style.VerticalAlignment = $VerticalAlignment   }

            if ($BackgroundColor) {
                $Range.Style.Fill.PatternType = $BackgroundPattern
                $Range.Style.Fill.BackgroundColor.SetColor($BackgroundColor)
                if ($PatternColor) {
                    $range.Style.Fill.PatternColor.SetColor( $PatternColor)
                }
            }

            if ($Height) {
                if ($Range -is [OfficeOpenXml.ExcelRow]   ) {$Range.Height = $Height }
                elseif ($Range -is [OfficeOpenXml.ExcelRange] ) {
                    ($range.Start.Row)..($range.Start.Row + $range.Rows) |
                        ForEach-Object {$ws.Row($_).Height = $Height }
                }
                else {Write-Warning -Message ("Can set the height of a row or a range but not a {0} object" -f ($Range.GetType().name)) }
            }
            if ($AutoFit) {
                if ($Range -is [OfficeOpenXml.ExcelColumn]) {$Range.AutoFit() }
                elseif ($Range -is [OfficeOpenXml.ExcelRange] ) {$Range.AutoFitColumns() }
                else {Write-Warning -Message ("Can autofit a column or a range but not a {0} object" -f ($Range.GetType().name)) }

            }
            elseif ($Width) {
                if ($Range -is [OfficeOpenXml.ExcelColumn]) {$Range.Width = $Width}
                elseif ($Range -is [OfficeOpenXml.ExcelRange] ) {
                    ($range.Start.Column)..($range.Start.Column + $range.Columns) |
                        ForEach-Object {$ws.Column($_).Width = $Width}
                }
                else {Write-Warning -Message ("Can set the width of a column or a range but not a {0} object" -f ($Range.GetType().name)) }
            }
            if ($Hidden) {
                if ($Range -is [OfficeOpenXml.ExcelRow] -or
                    $Range -is [OfficeOpenXml.ExcelColumn]  ) {$Range.Hidden = $True}
                else {Write-Warning -Message ("Can hide a row or a column but not a {0} object" -f ($Range.GetType().name)) }
            }
            if ($Indent) {$Range.Style.Indent = $Indent }
        }
    }
}
