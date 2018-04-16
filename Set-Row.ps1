Function Set-Row {
<#
.Synopsis
    Fills values into a row in a Excel spreadsheet
.Description
    Set-Row accepts either a Worksheet object or an Excel package object returned by Export-Excel and the name of a sheet,
    and inserts the chosen contents into a row of the sheet.
    The contents can be a constant "42" , a formula or a script block which is converted into a constant or formula.
    The first cell of the row can optional be given a heading.
.Example
    Set-row -Worksheet $ws -Heading Total -Value {"=sum($columnName`2:$columnName$endrow)" }

    $Ws contains a worksheet object, and no Row number is specified so Set-Row will select the next row after the end of the data in the sheet
    The first cell will contain "Total", and each other cell will contain
        =Sum(xx2:xx99)  - where xx is the column name, and 99 is the last row of data.
        Note the use of `2 to Prevent 2 becoming part of the variable "ColumnName"
    The script block can use $row, $column, $ColumnName, $startRow/Column $endRow/Column


#>
[cmdletbinding()]
    Param (
        #An Excel package object - e.g. from Export-Excel -passthru - requires a sheet name
        [Parameter(ParameterSetName="Package",Mandatory=$true)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        #the name  to update in the package
        [Parameter(ParameterSetName="Package")]
        $Worksheetname = "Sheet1",
        #A worksheet object
        [Parameter(ParameterSetName="sheet",Mandatory=$true)]
        [OfficeOpenXml.Excelworksheet]
        $Worksheet,
        #Row to fill right - first row is 1. 0 will be interpreted as first unused row
        $Row = 0 ,
        #Position in the row to start from
        [Int]$StartColumn,
        #value, formula or script block for to fill in. Script block can use $row, $column [number], $ColumnName [letter(s)], $startRow, $startColumn, $endRow, $endColumn
        [parameter(Mandatory=$true)]
        $Value,
        #Optional Row heading
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
        #Set cells to a fixed hieght
        [float]$Height,
        [switch]$PassThru
    )

    #if we were passed a package object and a worksheet name , get the worksheet.
    if ($ExcelPackage)     {$Worksheet   = $ExcelPackage.Workbook.worksheets[$Worksheetname] }

    #In a script block to build a formula, we may want any of corners or the columnname,
    #if row and start column aren't specified assume first unused row, and first column
    if (-not $StartColumn) {$StartColumn = $Worksheet.Dimension.Start.Column    }
    $startRow                            = $Worksheet.Dimension.Start.Row   + 1
    $endColumn                           = $Worksheet.Dimension.End.Column
    $endRow                              = $Worksheet.Dimension.End.Row
    if ($Row  -lt 2 )      {$Row         = $endRow + 1 }

    Write-Verbose -Message "Updating Row $Row"
    #Add a row label
    if      ($Heading)                   {
                                           $Worksheet.Cells[$Row, $StartColumn].Value = $Heading
                                           $StartColumn ++
    }
    #Fill in the data
    if      ($value) {foreach ($column in ($StartColumn..$EndColumn)) {
        #We might want the column name in a script block
        $ColumnName = [OfficeOpenXml.ExcelCellAddress]::new(1,$column).Address -replace "1",""
        if  ($Value -is [scriptblock] ) {
             #re-create the script block otherwise variables from this function are out of scope.
             $cellData = & ([scriptblock]::create( $Value ))
             Write-Verbose -Message $cellData
        }
        else{$cellData = $Value}
        if  ($cellData -match "^=")      { $Worksheet.Cells[$Row, $column].Formula                    = $cellData           }
        else                             { $Worksheet.Cells[$Row, $Column].Value                      = $cellData           }
        if  ($cellData -is [datetime])   { $Worksheet.Cells[$Row, $Column].Style.Numberformat.Format  = 'm/d/yy h:mm'       }
    }}
    #region Apply formatting
    if      ($Underline)                 {
                                           $worksheet.row(  $Row  ).Style.Font.UnderLine              = $true
                                           $worksheet.row(  $Row  ).Style.Font.UnderLineType          = $UnderLineType
    }
    if      ($Bold)                      { $worksheet.row(  $Row  ).Style.Font.Bold                   = $true               }
    if      ($Italic)                    { $worksheet.row(  $Row  ).Style.Font.Italic                 = $true               }
    if      ($StrikeThru)                { $worksheet.row(  $Row  ).Style.Font.Strike                 = $true               }
    if      ($FontShift)                 { $worksheet.row(  $Row  ).Style.Font.VerticalAlign          = $FontShift          }
    if      ($NumberFormat)              { $worksheet.row(  $Row  ).Style.Numberformat.Format         = $NumberFormat       }
    if      ($TextRotation)              { $worksheet.row(  $Row  ).Style.TextRotation                = $TextRotation       }
    if      ($WrapText)                  { $worksheet.row(  $Row  ).Style.WrapText                    = $true               }
    if      ($HorizontalAlignment)       { $worksheet.row(  $Row  ).Style.HorizontalAlignment         = $HorizontalAlignment}
    if      ($VerticalAlignment)         { $worksheet.row(  $Row  ).Style.VerticalAlignment           = $VerticalAlignment  }
    if      ($Height)                    { $worksheet.row(  $Row  ).Height                            = $Height             }
    if      ($FontColor)                 { $worksheet.row(  $Row  ).Style.Font.Color.SetColor(          $FontColor        ) }
    if      ($BorderAround)               { $worksheet.row(  $Row  ).Style.Border.BorderAround(          $BorderAround     ) }
    if      ($BackgroundColor)           {
                                           $worksheet.row(  $Row  ).Style.Fill.PatternType            = $BackgroundPattern
                                           $worksheet.row(  $Row  ).Style.Fill.BackgroundColor.SetColor($BackgroundColor  )
         if ($PatternColor)              { $worksheet.row(  $Row  ).Style.Fill.PatternColor.SetColor(   $PatternColor     ) }
    }
    #endregion
    #return the new data if -passthru was specified.
    if ($passThru) {$Worksheet.Row($Row)}
}