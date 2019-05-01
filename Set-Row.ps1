Function  Set-ExcelRow {
    <#
      .Synopsis
        Fills values into a [new] row in an Excel spreadsheet, and sets row formats.
      .Description
        Set-ExcelRow accepts either a Worksheet object or an ExcelPackage object
        returned by Export-Excel and the name of a sheet, and inserts the chosen
        contents into a row of the sheet. The contents can be a constant,
        like "42", a formula or a script block which is converted into a
        constant or a formula.
        The first cell of the row can optionally be given a heading.
      .Example
        Set-ExcelRow -Worksheet $ws -Heading Total -Value {"=sum($columnName`2:$columnName$endrow)" }

        $Ws contains a worksheet object, and no Row number is specified so
        Set-ExcelRow will select the next row after the endof the data in
        the sheet. The first cell in the row will contain "Total", and
        each of the other cells will contain
            =Sum(xx2:xx99)
        where xx is the column name, and 99 is the last row of data.
        Note the use of `2 to Prevent 2 becoming part of the variable "ColumnName"
        The script block can use $Worksheet, $Row, $Column (number),
        $ColumnName (letter), $StartRow/Column and $EndRow/Column.
      .Example
        Set-ExcelRow -Worksheet $ws -Heading Total -HeadingBold -Value {"=sum($columnName`2:$columnName$endrow)" } -NumberFormat 'Currency' -StartColumn 2 -Bold -BorderTop Double -BorderBottom Thin

        This builds on the previous example, but this time the label "Total"
        appears in column 2 and the formula fills from column 3 onwards.
        The formula and heading are set in bold face, and the formula is
        formatted for the local currency, and given a double line border
        above and single line border below.
    #>
    [cmdletbinding()]
    [Alias("Set-Row")]
    [OutputType([OfficeOpenXml.ExcelRow],[String])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',Justification='Does not change system state')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification="Variables created for script block which may be passed as a parameter, but not used in the script")]
    Param (
        #An Excel package object - e.g. from Export-Excel -PassThru - requires a sheet name.
        [Parameter(ParameterSetName="Package",Mandatory=$true)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        #The name of the sheet to update in the package.
        [Parameter(ParameterSetName="Package")]
        $Worksheetname = "Sheet1",
        #A worksheet object instead of passing a name and package.
        [Parameter(ParameterSetName="Sheet",Mandatory=$true)]
        [OfficeOpenXml.Excelworksheet] $Worksheet,
        #Row to fill right - first row is 1. 0 will be interpreted as first unused row.
        [Parameter(ValueFromPipeline = $true)]
        $Row = 0 ,
        #Position in the row to start from.
        [int]$StartColumn,
        #Value, Formula or ScriptBlock to fill in. A ScriptBlock can use $worksheet,  $row, $Column [number], $ColumnName [letter(s)], $startRow, $startColumn, $endRow, $endColumn.
        $Value,
        #Optional row-heading.
        $Heading ,
        #Set the heading in bold type.
        [Switch]$HeadingBold,
        #Change the font-size of the heading.
        [Int]$HeadingSize ,
        #Number format to apply to cells e.g. "dd/MM/yyyy HH:mm", "£#,##0.00;[Red]-£#,##0.00", "0.00%" , "##/##" , "0.0E+0" etc.
        [Alias("NFormat")]
        $NumberFormat,
        #Style of border to draw around the row.
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
        #Color for the text - if none specified it will be left as it it is.
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
        #Subscript or Superscript (or none).
        [OfficeOpenXml.Style.ExcelVerticalAlignmentFont]$FontShift,
        #Font to use - Excel defaults to Calibri.
        [String]$FontName,
        #Point size for the text.
        [float]$FontSize,
        #Change background color.
        $BackgroundColor,
        #Background pattern - solid by default.
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
        #Set cells to a fixed height.
        [float]$Height,
        #Hide the row.
        [Switch]$Hide,
        #If sepecified, returns the range of cells which were affected.
        [Switch]$ReturnRange,
        #If Specified, return a row object to allow further work to be done.
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
        #In a script block to build a formula, we may want any of corners or the columnname,
        #if row and start column aren't specified assume first unused row, and first column
        if (-not $StartColumn) {$StartColumn = $Worksheet.Dimension.Start.Column    }
        $startRow                            = $Worksheet.Dimension.Start.Row   + 1
        $endColumn                           = $Worksheet.Dimension.End.Column
        $endRow                              = $Worksheet.Dimension.End.Row
    }
    process {
        if ($null -eq $workSheet.Dimension) {Write-Warning "Can't format an empty worksheet."; return}
        if      ($Row  -eq 0 ) {$Row         = $endRow + 1 }
        Write-Verbose -Message "Updating Row $Row"
        #Add a row label
        if      ($Heading)     {
            $Worksheet.Cells[$Row, $StartColumn].Value = $Heading
            if ($HeadingBold) {$Worksheet.Cells[$Row, $StartColumn].Style.Font.Bold = $true}
            if ($HeadingSize) {$Worksheet.Cells[$Row, $StartColumn].Style.Font.Size = $HeadingSize}
            $StartColumn ++
        }
        #Fill in the data
        if      ($PSBoundParameters.ContainsKey('Value')) {foreach ($column in ($StartColumn..$endColumn)) {
            #We might want the column name in a script block
            $columnName = (New-Object -TypeName OfficeOpenXml.ExcelCellAddress @(1,$column)).Address -replace "1",""
            if  ($Value -is [scriptblock] ) {
                #re-create the script block otherwise variables from this function are out of scope.
                $cellData = & ([scriptblock]::create( $Value ))
                if ($null -eq $cellData) {Write-Verbose -Message "Script block evaluates to null."}
                else                     {Write-Verbose -Message "Script block evaluates to '$cellData'"}
            }
            else{$cellData = $Value}
            if  ($cellData -match "^=")      { $Worksheet.Cells[$Row, $column].Formula                    = ($cellData -replace '^=','') } #EPPlus likes formulas with no = sign; Excel doesn't care
            elseif ( [System.Uri]::IsWellFormedUriString($cellData , [System.UriKind]::Absolute)) {
                # Save a hyperlink : internal links can be in the form xl://sheet!E419 (use A1 as goto sheet), or xl://RangeName
                if ($cellData -match "^xl://internal/") {
                    $referenceAddress = $cellData -replace "^xl://internal/" , ""
                    $display          = $referenceAddress -replace "!A1$"    , ""
                    $h = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $referenceAddress , $display
                    $Worksheet.Cells[$Row, $Column].HyperLink = $h
                }
                else {$Worksheet.Cells[$Row, $Column].HyperLink = $cellData }
                $Worksheet.Cells[$Row, $Column].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
                $Worksheet.Cells[$Row, $Column].Style.Font.UnderLine = $true
            }
            else                             { $Worksheet.Cells[$Row, $column].Value                      = $cellData                    }
            if  ($cellData -is [datetime])   { $Worksheet.Cells[$Row, $column].Style.Numberformat.Format  = 'm/d/yy h:mm'                } #This is not a custom format, but a preset recognized as date and localized.
            if  ($cellData -is [timespan])   { $Worksheet.Cells[$Row, $Column].Style.Numberformat.Format  = '[h]:mm:ss'                  }
        }}
        #region Apply formatting
        $params = @{}
        foreach ($p in @('Underline','Bold','Italic','StrikeThru', 'FontName', 'FontSize', 'FontShift','NumberFormat','TextRotation',
                        'WrapText', 'HorizontalAlignment','VerticalAlignment', 'Height', 'FontColor'
                        'BorderAround', 'BorderBottom', 'BorderTop', 'BorderLeft', 'BorderRight', 'BorderColor',
                        'BackgroundColor', 'BackgroundPattern', 'PatternColor')) {
            if ($PSBoundParameters.ContainsKey($p)) {$params[$p] = $PSBoundParameters[$p]}
        }
        if ($params.Count) {
            $theRange = New-Object -TypeName OfficeOpenXml.ExcelAddress @($Row, $StartColumn, $Row, $endColumn)
            Set-ExcelRange -WorkSheet $Worksheet -Range $theRange @params
        }
        #endregion
        if ($PSBoundParameters.ContainsKey('Hide')) {$workSheet.Row($Row).Hidden = [bool]$Hide}
        #return the new data if -passthru was specified.
        if     ($passThru)     {$Worksheet.Row($Row)}
        elseif ($ReturnRange)  {$theRange}
    }
}