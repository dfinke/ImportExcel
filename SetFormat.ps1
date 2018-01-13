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
        [Parameter(ValueFromPipeline = $true,ParameterSetName="Address",Mandatory=$True)]
        $Address ,
        #The worksheet where the format is to be applied
        [Parameter(ParameterSetName="SheetAndRange",Mandatory=$True)]
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet ,
        #The area of the worksheet where the format is to be applied
        [Parameter(ParameterSetName="SheetAndRange",Mandatory=$True)]
        [OfficeOpenXml.ExcelAddress]$Range,
        #Number format to apply to cells e.g. "dd/MM/yyyy HH:mm", "Â£#,##0.00;[Red]-Â£#,##0.00", "0.00%" , "##/##" , "0.0E+0" etc
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
        [Alias("AutoFit")]
        [Switch]$AutoSize,
        #Set cells to a fixed width (columns or ranges only), ignored if Autosize is specified
        [float]$Width,
        #Set cells to a fixed hieght  (rows or ranges only)
        [float]$Height,
        #Hide a row or column  (not a range)
        [switch]$Hidden
    )
    begin {
        #Allow Set-Format to take Worksheet and range parameters (like Add Contitional formatting) -  convert them to an address
        if ($WorkSheet -and $Range) {$Address = $WorkSheet.Cells[$Range] }
    }

    process {
        if   ($Address -is [Array])  {
            [void]$PSBoundParameters.Remove("Address")
            $Address | Set-Format @PSBoundParameters
        }
        else {
            if ($ResetFont) {
                $Address.Style.Font.Color.SetColor("Black")
                $Address.Style.Font.Bold = $false
                $Address.Style.Font.Italic = $false
                $Address.Style.Font.UnderLine = $false
                $Address.Style.Font.Strike = $false
            }
            if ($Underline) {
                $Address.Style.Font.UnderLine = $true
                $Address.Style.Font.UnderLineType = $UnderLineType
            }
            if ($Bold) {$Address.Style.Font.Bold = $true                }
            if ($Italic) {$Address.Style.Font.Italic = $true                }
            if ($StrikeThru) {$Address.Style.Font.Strike = $true                }
            if ($FontShift) {$Address.Style.Font.VerticalAlign = $FontShift           }
            if ($FontColor) {$Address.Style.Font.Color.SetColor( $FontColor    )      }
            if ($BorderAround) {$Address.Style.Border.BorderAround( $BorderAround )      }
            if ($NumberFormat) {$Address.Style.Numberformat.Format = $NumberFormat        }
            if ($TextRotation) {$Address.Style.TextRotation = $TextRotation        }
            if ($WrapText) {$Address.Style.WrapText = $true                }
            if ($HorizontalAlignment) {$Address.Style.HorizontalAlignment = $HorizontalAlignment }
            if ($VerticalAlignment) {$Address.Style.VerticalAlignment = $VerticalAlignment   }

            if ($BackgroundColor) {
                $Address.Style.Fill.PatternType = $BackgroundPattern
                $Address.Style.Fill.BackgroundColor.SetColor($BackgroundColor)
                if ($PatternColor) {
                    $Address.Style.Fill.PatternColor.SetColor( $PatternColor)
                }
            }

            if ($Height) {
                if ($Address -is [OfficeOpenXml.ExcelRow]   ) {$Address.Height = $Height }
                elseif ($Address -is [OfficeOpenXml.ExcelRange] ) {
                    ($Address.Start.Row)..($Address.Start.Row + $Address.Rows) |
                        ForEach-Object {$ws.Row($_).Height = $Height }
                }
                else {Write-Warning -Message ("Can set the height of a row or a range but not a {0} object" -f ($Address.GetType().name)) }
            }
            if ($Autosize) {
                if ($Address -is [OfficeOpenXml.ExcelColumn]) {$Address.AutoFit() }
                elseif ($Address -is [OfficeOpenXml.ExcelRange] ) {$Address.AutoFitColumns() }
                else {Write-Warning -Message ("Can autofit a column or a range but not a {0} object" -f ($Address.GetType().name)) }

            }
            elseif ($Width) {
                if ($Address -is [OfficeOpenXml.ExcelColumn]) {$Address.Width = $Width}
                elseif ($Address -is [OfficeOpenXml.ExcelRange] ) {
                    ($Address.Start.Column)..($Address.Start.Column + $Address.Columns) |
                        ForEach-Object {$ws.Column($_).Width = $Width}
                }
                else {Write-Warning -Message ("Can set the width of a column or a range but not a {0} object" -f ($Address.GetType().name)) }
            }
            if ($Hidden) {
                if ($Address -is [OfficeOpenXml.ExcelRow] -or
                    $Address -is [OfficeOpenXml.ExcelColumn]  ) {$Address.Hidden = $True}
                else {Write-Warning -Message ("Can hide a row or a column but not a {0} object" -f ($Address.GetType().name)) }
            }
        }
    }
}