Function Set-ExcelColumn {
    <#
      .SYNOPSIS
        Adds or modifies a column in an Excel worksheet, filling values, setting formatting and/or creating named ranges.
      .DESCRIPTION
        Set-ExcelColumn can take a value which is either a string containing a
        value or formula or a scriptblock which evaluates to a string,
        and optionally a column number and fills that value down the column.
        A column heading can be specified, and the column can be made a named range.
        The column can be formatted in the same operation.
      .EXAMPLE
        Set-ExcelColumn -Worksheet $ws -Column 5 -NumberFormat 'Currency'

        $ws contains a worksheet object - and column "E" is set to use the
        local currency format. Intelisense will complete the names of predefined
        number formats. You can see how currency is interpreted on the
        local computer with the command
        Expand-NumberFormat currency
      .EXAMPLE
        Set-ExcelColumn -Worksheet $ws -Heading "WinsToFastLaps"  -Value {"=E$row/C$row"} -Column 7 -AutoSize -AutoNameRange

        Here, $WS already contains a worksheet which holds counts of races won
        and fastest laps recorded by racing drivers (in columns C and E).
        Set-ExcelColumn specifies that Column 7 should have a heading of
        "WinsToFastLaps" and the data cells should contain =E2/C2 , =E3/C3 etc
        the new data cells should become a named range, which will also be
        named "WinsToFastLaps" and the column width will be set automatically.
      .EXAMPLE
        Set-ExcelColumn -Worksheet $ws -Heading "Link" -Value {"https://en.wikipedia.org" + $worksheet.cells["B$Row"].value  }  -AutoSize

        In this example, the worksheet in $ws has partial links to Wikipedia
        pages in column B. The -Value parameter is a script block which
        outputs a string beginning "https..." and ending with the value of
        the cell at column B in the current row.
        When given a valid URI,  Set-ExcelColumn makes it a hyperlink.
        The column will be autosized to fit the links.
      .EXAMPLE
        4..6 | Set-ExcelColumn -Worksheet $ws -AutoNameRange

        Again $ws contains a worksheet. Here columns 4 to 6 are made into
        named ranges, row 1 is used for the range name
        and the rest of the column becomes the range.
    #>
    [cmdletbinding()]
    [Alias("Set-Column")]
    [OutputType([OfficeOpenXml.ExcelColumn],[String])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',Justification='Does not change system state')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification="Variables created for script block which may be passed as a parameter, but not used in the script")]
    Param (
        #If specifying the worksheet by name, the ExcelPackage object which contains the worksheet also needs to be passed.
        [Parameter(ParameterSetName="Package",Mandatory=$true)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        #The sheet to update can be given as a name or an Excel Worksheet object - this sets it by name.
        [Parameter(ParameterSetName="Package")]
        [String]$Worksheetname = "Sheet1",
        #This passes the worksheet object instead of passing a sheet name and an Excelpackage object.
        [Parameter(ParameterSetName="sheet",Mandatory=$true)]
        [OfficeOpenXml.ExcelWorksheet]$Worksheet,
        #Column to fill down - the first column is 1. 0 will be interpreted as first empty column.
        [Parameter(ValueFromPipeline=$true)]
        [ValidateRange(0,16384)]
        $Column = 0 ,
        #First row to fill data in.
        [ValidateRange(1,1048576)]
        [Int]$StartRow ,
        #A value, formula or scriptblock to fill in. A script block can use $worksheet, $row, $column [number], $columnName [letter(s)], $startRow, $startColumn, $endRow, $endColumn.
        $Value ,
        #Optional column heading.
        $Heading ,
        #Number format to apply to cells e.g. "dd/MM/yyyy HH:mm", "£#,##0.00;[Red]-£#,##0.00", "0.00%" , "##/##" , "0.0E+0" etc.
        [Alias("NFormat")]
        $NumberFormat,
        #Style of border to draw around the row.
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderAround,
        #Colour for the text - if none specified it will be left as it it is.
        $FontColor,
        #Make text bold; use -Bold:$false to remove bold.
        [Switch]$Bold,
        #Make text italic;  use -Italic:$false to remove italic.
        [Switch]$Italic,
        #Underline the text using the underline style in -UnderlineType;  use -Underline:$false to remove underlining.
        [Switch]$Underline,
        #Specifies whether underlining should be single or double, normal or accounting mode. The default is "Single".
        [OfficeOpenXml.Style.ExcelUnderLineType]$UnderLineType = [OfficeOpenXml.Style.ExcelUnderLineType]::Single,
        #Strike through text; use -StrikeThru:$false to remove strike through.
        [Switch]$StrikeThru,
        #Subscript or Superscript (or None).
        [OfficeOpenXml.Style.ExcelVerticalAlignmentFont]$FontShift,
        #Font to use - Excel defaults to Calibri.
        [String]$FontName,
        #Point size for the text.
        [float]$FontSize,
        #Change background color.
        $BackgroundColor,
        #Background pattern - "Solid" by default.
        [OfficeOpenXml.Style.ExcelFillStyle]$BackgroundPattern = [OfficeOpenXml.Style.ExcelFillStyle]::Solid ,
        #Secondary color for background pattern.
        [Alias("PatternColour")]
        $PatternColor,
        #Turn on Text-Wrapping; use -WrapText:$false to turn off wrapping.
        [Switch]$WrapText,
        #Position cell contents to Left, Right, Center etc. Default is "General".
        [OfficeOpenXml.Style.ExcelHorizontalAlignment]$HorizontalAlignment,
        #Position cell contents to Top, Bottom or Center.
        [OfficeOpenXml.Style.ExcelVerticalAlignment]$VerticalAlignment,
        #Degrees to rotate text. Up to +90 for anti-clockwise ("upwards"), or to -90 for clockwise.
        [ValidateRange(-90, 90)]
        [int]$TextRotation ,
        #Attempt to auto-fit cells to the width their contents.
        [Alias("AutoFit")]
        [Switch]$AutoSize,
        #Set cells to a fixed width, ignored if -AutoSize is specified.
        [float]$Width,
        #Set the inserted data to be a named range.
        [Switch]$AutoNameRange,
        #Hide the column.
        [Switch]$Hide,
        #If specified, returns the range of cells which were affected.
        [Switch]$Specified,
        #If specified, return an object representing the Column, to allow further work to be done on it.
        [Switch]$PassThru
    )

    begin {
        #if we were passed a package object and a worksheet name , get the worksheet.
        if ($ExcelPackage)  {
            if ($ExcelPackage.Workbook.Worksheets.Name -notcontains $Worksheetname) {
                throw "The Workbook does not contain a sheet named '$Worksheetname'"
            }
            else {$Worksheet   = $ExcelPackage.Workbook.Worksheets[$Worksheetname] }
        }

        #In a script block to build a formula, we may want any of corners or the column name,
        #if Column and Startrow aren't specified, assume first unused column, and first row
        if (-not $StartRow)   {$startRow   = $Worksheet.Dimension.Start.Row    }
        $startColumn                       = $Worksheet.Dimension.Start.Column
        $endColumn                         = $Worksheet.Dimension.End.Column
        $endRow                            = $Worksheet.Dimension.End.Row
    }
    process {
        if ($null -eq $workSheet.Dimension) {Write-Warning "Can't format an empty worksheet."; return}
        if ($Column  -eq 0 )  {$Column     = $endColumn    + 1 }
        $columnName = (New-Object 'OfficeOpenXml.ExcelCellAddress' @(1, $column)).Address -replace "1",""
        Write-Verbose -Message "Updating Column $columnName"
        #If there is a heading, insert it and use it as the name for a range (if we're creating one)
        if      ($PSBoundParameters.ContainsKey('Heading'))                 {
                 $Worksheet.Cells[$StartRow, $Column].Value = $Heading
                 $StartRow ++
                 if ($AutoNameRange) {
                    Add-ExcelName -Range $Worksheet.Cells[$StartRow, $Column, $endRow, $Column] -RangeName $Heading
                }
        }
        elseif ($AutoNameRange) {
            Add-ExcelName -Range $Worksheet.Cells[($StartRow+1), $Column, $endRow, $Column] -RangeName $Worksheet.Cells[$StartRow, $Column].Value
        }

        #Fill in the data -it can be zero null or and empty string.
        if      ($PSBoundParameters.ContainsKey('Value')) { foreach ($row in ($StartRow..$endRow)) {
            if  ($Value -is [scriptblock]) { #re-create the script block otherwise variables from this function are out of scope.
                 $cellData = & ([scriptblock]::create( $Value ))
                 if ($null -eq $cellData) {Write-Verbose -Message "Script block evaluates to null."}
                 else                     {Write-Verbose -Message "Script block evaluates to '$cellData'"}
            }
            else                           { $cellData = $Value}
            if  ($cellData -match "^=")    { $Worksheet.Cells[$Row, $Column].Formula = ($cellData -replace '^=','') } #EPPlus likes formulas with no = sign; Excel doesn't care
            elseif ( [System.Uri]::IsWellFormedUriString($cellData , [System.UriKind]::Absolute)) {
                # Save a hyperlink : internal links can be in the form xl://sheet!E419 (use A1 as goto sheet), or xl://RangeName
                if ($cellData -match "^xl://internal/") {
                      $referenceAddress = $cellData -replace "^xl://internal/" , ""
                      $display          = $referenceAddress -replace "!A1$"    , ""
                      $h = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $referenceAddress , $display
                      $Worksheet.Cells[$Row, $Column].HyperLink = $h
                }
                else {$Worksheet.Cells[$Row, $Column].HyperLink      = $cellData }
                $Worksheet.Cells[$Row, $Column].Style.Font.UnderLine = $true
                $Worksheet.Cells[$Row, $Column].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
            }
            else                           { $Worksheet.Cells[$Row, $Column].Value                     = $cellData     }
            if  ($cellData -is [datetime]) { $Worksheet.Cells[$Row, $Column].Style.Numberformat.Format = 'm/d/yy h:mm' } # This is not a custom format, but a preset recognized as date and localized.
            if  ($cellData -is [timespan]) { $Worksheet.Cells[$Row, $Column].Style.Numberformat.Format = '[h]:mm:ss'   }
        }}

        #region Apply formatting
        $params = @{}
        foreach ($p in @('Underline','Bold','Italic','StrikeThru', 'FontName', 'FontSize','FontShift','NumberFormat','TextRotation',
                         'WrapText', 'HorizontalAlignment','VerticalAlignment', 'Autosize', 'Width', 'FontColor'
                         'BorderAround', 'BackgroundColor', 'BackgroundPattern', 'PatternColor')) {
            if ($PSBoundParameters.ContainsKey($p)) {$params[$p] = $PSBoundParameters[$p]}
        }
        if ($params.Count) {
            $theRange =   "$columnName$StartRow`:$columnName$endRow"
            Set-ExcelRange -WorkSheet $Worksheet -Range $theRange @params
        }
        #endregion
        if      ($PSBoundParameters.ContainsKey('Hide')) {$workSheet.Column($Column).Hidden = [bool]$Hide}
        #return the new data if -passthru was specified.
        if      ($PassThru)                 { $Worksheet.Column($Column)}
        elseif  ($ReturnRange)              { $theRange}
    }
}