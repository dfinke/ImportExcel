function  Set-ExcelRow {
    [CmdletBinding()]
    [Alias("Set-Row")]
    [OutputType([OfficeOpenXml.ExcelRow],[String])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingfunctions', '',Justification='Does not change system state')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification="Variables created for script block which may be passed as a parameter, but not used in the script")]
    param(
        [Parameter(ParameterSetName="Package",Mandatory=$true)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        [Parameter(ParameterSetName="Package")]
        $WorksheetName = "Sheet1",
        [Parameter(ParameterSetName="Sheet",Mandatory=$true)]
        [OfficeOpenXml.Excelworksheet] $Worksheet,
        [Parameter(ValueFromPipeline = $true)]
        $Row = 0 ,
        [int]$StartColumn,
        $Value,
        $Heading ,
        [Switch]$HeadingBold,
        [Int]$HeadingSize ,
        [Alias("NFormat")]
        $NumberFormat,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderAround,
        $BorderColor=[System.Drawing.Color]::Black,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderBottom,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderTop,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderLeft,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderRight,
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
        [float]$Height,
        [Alias('Hidden')]
        [Switch]$Hide,
        [Switch]$ReturnRange,
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
        #In a script block to build a formula, we may want any of corners or the columnname,
        #if row and start column aren't specified assume first unused row, and first column
        if (-not $StartColumn) {$StartColumn = $Worksheet.Dimension.Start.Column    }
        $startRow                            = $Worksheet.Dimension.Start.Row   + 1
        $endColumn                           = $Worksheet.Dimension.End.Column
        $endRow                              = $Worksheet.Dimension.End.Row
    }
    process {
        if ($null -eq $Worksheet.Dimension) {Write-Warning "Can't format an empty worksheet."; return}
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
            Set-ExcelRange -Worksheet $Worksheet -Range $theRange @params
        }
        #endregion
        if ($PSBoundParameters.ContainsKey('Hide')) {$Worksheet.Row($Row).Hidden = [bool]$Hide}
        #return the new data if -passthru was specified.
        if     ($passThru)     {$Worksheet.Row($Row)}
        elseif ($ReturnRange)  {$theRange}
    }
}