Function Set-Column {
<#
    .SYNOPSIS
        Adds a column to the existing data area in an Excel sheet, fills values and sets formatting
    .DESCRIPTION
        Set-Column takes a value which is either string containing a value or formula or a scriptblock
        which evaluates to a string, and optionally a column number and fills that value down the column.
        A column name can be specified and the new column can be made a named range.
        The column can be formatted.
    .Example
        C:> Set-Column -Worksheet $ws -Heading "WinsToFastLaps"  -Value {"=E$row/C$row"} -Column 7 -AutoSize -AutoNameRange
        Here $WS already contains a worksheet which contains counts of races won and fastest laps recorded by racing drivers (in columns C and E)
        Set-Column specifies that Column 7 should have a heading of "WinsToFastLaps" and the data cells should contain =E2/C2 , =E3/C3
        the data celss should become a named range, which will also be "WinsToFastLaps" the column width will be set automatically

#>
[cmdletbinding()]
    Param (
        [Parameter(ParameterSetName="Package",Mandatory=$true)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        #Sheet to update
        [Parameter(ParameterSetName="Package")]
        $Worksheetname = "Sheet1",
        [Parameter(ParameterSetName="sheet",Mandatory=$true)]
        [OfficeOpenXml.ExcelWorksheet]
        $Worksheet,
        #Column to fill down - first column is 1. 0 will be interpreted as first unused column
        $Column = 0 ,
        [Int]$StartRow ,
        #value, formula or script block for to fill in. Script block can use $row, $column [number], $ColumnName [letter(s)], $startRow, $startColumn, $endRow, $endColumn
        [parameter(Mandatory=$true)]
        $Value ,
        #Optional column heading
        $Heading ,
        #Number format to apply to cells e.g. "dd/MM/yyyy HH:mm", "Â£#,##0.00;[Red]-Â£#,##0.00", "0.00%" , "##/##" , "0.0E+0" etc
        [Alias("NFormat")]
        $NumberFormat,
        #Style of border to draw around the row
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderAround,
        #Colour for the text - if none specified it will be left as it it is
        [System.Drawing.Color]$FontColor,
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
        #Autofit cells to width
        [Alias("AutoFit")]
        [Switch]$AutoSize,
        #Set cells to a fixed width, ignored if Autosize is specified
        [float]$Width,
        #Set the inserted data to be a named range (ignored if header is not specified) d
        [Switch]$AutoNameRange,
        [switch]$PassThru
    )
    #if we were passed a package object and a worksheet name , get the worksheet.
    if ($ExcelPackage)   {$Worksheet   = $ExcelPackage.Workbook.Worksheets[$Worksheetname] }

    #In a script block to build a formula, we may want any of corners or the columnname,
    #if column and startrow aren't specified, assume first unused column, and first row
    if (-not $StartRow)   {$startRow   = $Worksheet.Dimension.Start.Row    }
    $StartColumn                       = $Worksheet.Dimension.Start.Column
    $endColumn                         = $Worksheet.Dimension.End.Column
    $endRow                            = $Worksheet.Dimension.End.Row
    if ($Column  -lt 2 )  {$Column     = $endColumn    + 1 }
    $ColumnName = [OfficeOpenXml.ExcelCellAddress]::new(1,$column).Address -replace "1",""

    Write-Verbose -Message "Updating Column $ColumnName"
    #If there is a heading, insert it and use it as the name for a range (if we're creating one)
    if      ($Heading)                 {
                                         $Worksheet.Cells[$StartRow, $Column].Value = $heading
                                         $startRow ++
      if    ($AutoNameRange)           { $Worksheet.Names.Add(  $heading, ($Worksheet.Cells[$startrow, $Column, $endRow, $Column]) ) | Out-Null }
    }
    #Fill in the data
    if      ($value)                   { foreach ($row in ($StartRow.. $endRow)) {
        if  ($Value -is [scriptblock]) { #re-create the script block otherwise variables from this function are out of scope.
             $cellData = & ([scriptblock]::create( $Value ))
             Write-Verbose  -Message     $cellData
        }
        else                           { $cellData = $Value}
        if  ($cellData -match "^=")    { $Worksheet.Cells[$Row, $Column].Formula                           = $cellData           }
        else                           { $Worksheet.Cells[$Row, $Column].Value                             = $cellData           }
        if  ($cellData -is [datetime]) { $Worksheet.Cells[$Row, $Column].Style.Numberformat.Format         = 'm/d/yy h:mm'       }
    }}
    #region Apply formatting
    if      ($Underline)               {
                                         $Worksheet.Column(     $Column).Style.Font.UnderLine              = $true
                                         $Worksheet.Column(     $Column).Style.Font.UnderLineType          = $UnderLineType
    }
    if      ($Bold)                    { $Worksheet.Column(     $Column).Style.Font.Bold                   = $true               }
    if      ($Italic)                  { $Worksheet.Column(     $Column).Style.Font.Italic                 = $true               }
    if      ($StrikeThru)              { $Worksheet.Column(     $Column).Style.Font.Strike                 = $true               }
    if      ($FontShift)               { $Worksheet.Column(     $Column).Style.Font.VerticalAlign          = $FontShift          }
    if      ($NumberFormat)            { $Worksheet.Column(     $Column).Style.Numberformat.Format         = $NumberFormat       }
    if      ($TextRotation)            { $Worksheet.Column(     $Column).Style.TextRotation                = $TextRotation       }
    if      ($WrapText)                { $Worksheet.Column(     $Column).Style.WrapText                    = $true               }
    if      ($HorizontalAlignment)     { $Worksheet.Column(     $Column).Style.HorizontalAlignment         = $HorizontalAlignment}
    if      ($VerticalAlignment)       { $Worksheet.Column(     $Column).Style.VerticalAlignment           = $VerticalAlignment  }
    if      ($FontColor)               { $Worksheet.Column(     $Column).Style.Font.Color.SetColor(          $FontColor        ) }
    if      ($BorderAround)             { $Worksheet.Column(     $Column).Style.Border.BorderAround(          $BorderAround     ) }
    if      ($BackgroundColor)         {
                                         $Worksheet.Column(     $Column).Style.Fill.PatternType            = $BackgroundPattern
                                         $Worksheet.Column(     $Column).Style.Fill.BackgroundColor.SetColor($BackgroundColor  )
         if ($PatternColor)            { $Worksheet.Column(     $Column).Style.Fill.PatternColor.SetColor(   $PatternColor     ) }
     }
     if     ($Autosize)                { $Worksheet.Column(     $Column).AutoFit()                                               }
     elseif ($Width)                   { $Worksheet.Column(     $Column).Width                             = $Width              }
     #endregion
    #return the new data if -passthru was specified.
    if     ($passThru)                 { $Worksheet.Column(     $Column)}
}