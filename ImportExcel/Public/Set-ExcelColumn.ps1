function Set-ExcelColumn {
    [CmdletBinding()]
    [Alias("Set-Column")]
    [OutputType([OfficeOpenXml.ExcelColumn],[String])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingfunctions', '',Justification='Does not change system state')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification="Variables created for script block which may be passed as a parameter, but not used in the script")]
    param(
        [Parameter(ParameterSetName="Package",Mandatory=$true)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        [Parameter(ParameterSetName="Package")]
        [String]$WorksheetName = "Sheet1",
        [Parameter(ParameterSetName="sheet",Mandatory=$true)]
        [OfficeOpenXml.ExcelWorksheet]$Worksheet,
        [Parameter(ValueFromPipeline=$true)]
        [ValidateRange(0,16384)]
        $Column = 0 ,
        [ValidateRange(1,1048576)]
        [Int]$StartRow ,
        $Value ,
        $Heading ,
        [Alias("NFormat")]
        $NumberFormat,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderAround,
        $FontColor,
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
        [Switch]$AutoNameRange,
        [Alias('Hidden')]
        [Switch]$Hide,
        [Switch]$Specified,
        [Switch]$PassThru
    )

    begin {
        #if we were passed a package object and a worksheet name , get the worksheet.
        if ($ExcelPackage)  {
            if ($ExcelPackage.Workbook.Worksheets.Name -notcontains $WorksheetName) {
                throw "The Workbook does not contain a sheet named '$WorksheetName'"
            }
            else {$Worksheet   = $ExcelPackage.Workbook.Worksheets[$WorksheetName] }
        }

        #In a script block to build a formula, we may want any of corners or the column name,
        #if Column and Startrow aren't specified, assume first unused column, and first row
        if (-not $StartRow)   {$startRow   = $Worksheet.Dimension.Start.Row    }
        $startColumn                       = $Worksheet.Dimension.Start.Column
        $endColumn                         = $Worksheet.Dimension.End.Column
        $endRow                            = $Worksheet.Dimension.End.Row
    }
    process {
        if ($null -eq $Worksheet.Dimension) {Write-Warning "Can't format an empty worksheet."; return}
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
        foreach ($p in @('Underline','UnderLineType','Bold','Italic','StrikeThru', 'FontName', 'FontSize','FontShift','NumberFormat','TextRotation',
                         'WrapText', 'HorizontalAlignment','VerticalAlignment', 'Autosize', 'Width', 'FontColor'
                         'BorderAround', 'BackgroundColor', 'BackgroundPattern', 'PatternColor')) {
            if ($PSBoundParameters.ContainsKey($p)) {$params[$p] = $PSBoundParameters[$p]}
        }
        if ($params.Count) {
            $theRange =   "$columnName$StartRow`:$columnName$endRow"
            Set-ExcelRange -Worksheet $Worksheet -Range $theRange @params
        }
        #endregion
        if      ($PSBoundParameters.ContainsKey('Hide')) {$Worksheet.Column($Column).Hidden = [bool]$Hide}
        #return the new data if -passthru was specified.
        if      ($PassThru)                 { $Worksheet.Column($Column)}
        elseif  ($ReturnRange)              { $theRange}
    }
}