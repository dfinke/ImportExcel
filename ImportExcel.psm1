﻿#region import everything we need
    Add-Type -Path "$($PSScriptRoot)\EPPlus.dll"
    . $PSScriptRoot\AddConditionalFormatting.ps1
    . $PSScriptRoot\Charting.ps1
    . $PSScriptRoot\ColorCompletion.ps1
    . $PSScriptRoot\ConvertExcelToImageFile.ps1
    . $PSScriptRoot\Compare-WorkSheet.ps1
    . $PSScriptRoot\ConvertFromExcelData.ps1
    . $PSScriptRoot\ConvertFromExcelToSQLInsert.ps1
    . $PSScriptRoot\ConvertToExcelXlsx.ps1
    . $PSScriptRoot\Copy-ExcelWorkSheet.ps1
    . $PSScriptRoot\Export-Excel.ps1
    . $PSScriptRoot\Export-ExcelSheet.ps1
    . $PSScriptRoot\Get-ExcelColumnName.ps1
    . $PSScriptRoot\Get-ExcelSheetInfo.ps1
    . $PSScriptRoot\Get-ExcelWorkbookInfo.ps1
    . $PSScriptRoot\Get-HtmlTable.ps1
    . $PSScriptRoot\Get-Range.ps1
    . $PSScriptRoot\Get-XYRange.ps1
    . $PSScriptRoot\Import-Html.ps1
    . $PSScriptRoot\InferData.ps1
    . $PSScriptRoot\Invoke-Sum.ps1
    . $PSScriptRoot\Join-WorkSheet.ps1
    . $PSScriptRoot\Merge-Worksheet.ps1
    . $PSScriptRoot\New-ConditionalFormattingIconSet.ps1
    . $PSScriptRoot\New-ConditionalText.ps1
    . $PSScriptRoot\New-ExcelChart.ps1
    . $PSScriptRoot\New-PSItem.ps1
    . $PSScriptRoot\Open-ExcelPackage.ps1
    . $PSScriptRoot\Pivot.ps1
    . $PSScriptRoot\Send-SQLDataToExcel.ps1
    . $PSScriptRoot\Set-CellStyle.ps1
    . $PSScriptRoot\Set-Column.ps1
    . $PSScriptRoot\Set-Row.ps1
    . $PSScriptRoot\SetFormat.ps1
    . $PSScriptRoot\TrackingUtils.ps1
    . $PSScriptRoot\Update-FirstObjectProperties.ps1


    New-Alias -Name Use-ExcelData -Value "ConvertFrom-ExcelData" -Force

    if ($PSVersionTable.PSVersion.Major -ge 5) {
        . $PSScriptRoot\Plot.ps1

        Function New-Plot {
            Param()

            [PSPlot]::new()
        }

    }
    else {
        Write-Warning 'PowerShell 5 is required for plot.ps1'
        Write-Warning 'PowerShell Excel is ready, except for that functionality'
    }
#endregion
function Import-Excel {
  <#
   .SYNOPSIS
       Create custom objects from the rows in an Excel worksheet.

   .DESCRIPTION
       The Import-Excel cmdlet creates custom objects from the rows in an Excel worksheet. Each row represents one object. All of this is possible without installing Microsoft Excel and by using the .NET library ‘EPPLus.dll’.

       By default, the property names of the objects are retrieved from the column headers. Because an object cannot have a blanc property name, only columns with column headers will be imported.

       If the default behavior is not desired and you want to import the complete worksheet ‘as is’, the parameter ‘-NoHeader’ can be used. In case you want to provide your own property names, you can use the parameter ‘-HeaderName’.

   .PARAMETER Path
       Specifies the path to the Excel file.

   .PARAMETER WorksheetName
       Specifies the name of the worksheet in the Excel workbook to import. By default, if no name is provided, the first worksheet will be imported.

   .PARAMETER DataOnly
       Import only rows and columns that contain data, empty rows and empty columns are not imported.

   .PARAMETER HeaderName
       Specifies custom property names to use, instead of the values defined in the column headers of the TopRow.

       In case you provide less header names than there is data in the worksheet, then only the data with a corresponding header name will be imported and the data without header name will be disregarded.

       In case you provide more header names than there is data in the worksheet, then all data will be imported and all objects will have all the property names you defined in the header names. As such, the last properties will be blanc as there is no data for them.

   .PARAMETER NoHeader
       Automatically generate property names (P1, P2, P3, ..) instead of the ones defined in the column headers of the TopRow.

       This switch is best used when you want to import the complete worksheet ‘as is’ and are not concerned with the property names.

   .PARAMETER StartRow
       The row from where we start to import data, all rows above the StartRow are disregarded. By default this is the first row.

       When the parameters ‘-NoHeader’ and ‘-HeaderName’ are not provided, this row will contain the column headers that will be used as property names. When one of both parameters are provided, the property names are automatically created and this row will be treated as a regular row containing data.

   .PARAMETER EndRow
       By default all rows up to the last cell in the sheet will be imported. If specified, import stops at this row.

   .PARAMETER StartColumn
        The number of the first column to read data from (1 by default).

   .PARAMETER EndColumn
        By default the import reads up to the last populated column, -EndColumn tells the import to stop at an earlier number.

   .PARAMETER Password
       Accepts a string that will be used to open a password protected Excel file.

   .EXAMPLE
       Import data from an Excel worksheet. One object is created for each row. The property names of the objects consist of the column names defined in the first row. In case a column doesn’t have a column header (usually in row 1 when ‘-StartRow’ is not used), then the unnamed columns will be skipped and the data in those columns will not be imported.

       ----------------------------------------------
       | File: Movies.xlsx     -      Sheet: Actors |
       ----------------------------------------------
       |           A           B            C       |
       |1     First Name                 Address    |
       |2     Chuck         Norris       California |
       |3     Jean-Claude   Vandamme     Brussels   |
       ----------------------------------------------

       PS C:\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Actors

       First Name: Chuck
       Address   : California

       First Name: Jean-Claude
       Address   : Brussels

       Notice that column 'B' is not imported because there's no value in cell 'B1' that can be used as property name for the objects.

   .EXAMPLE
       Import the complete Excel worksheet ‘as is’ by using the ‘-NoHeader’ switch. One object is created for each row. The property names of the objects will be automatically generated (P1, P2, P3, ..).

       ----------------------------------------------
       | File: Movies.xlsx     -      Sheet: Actors |
       ----------------------------------------------
       |           A           B            C       |
       |1     First Name                 Address    |
       |2     Chuck         Norris       California |
       |3     Jean-Claude   Vandamme     Brussels   |
       ----------------------------------------------

       PS C:\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Actors -NoHeader

       P1: First Name
       P2:
       P3: Address

       P1: Chuck
       P2: Norris
       P3: California

       P1: Jean-Claude
       P2: Vandamme
       P3: Brussels

       Notice that the column header (row 1) is imported as an object too.

    .EXAMPLE
       Import data from an Excel worksheet. One object is created for each row. The property names of the objects consist of the names defined in the parameter ‘-HeaderName’. The properties are named starting from the most left column (A) to the right. In case no value is present in one of the columns, that property will have an empty value.

       ----------------------------------------------------------
       | File: Movies.xlsx            -           Sheet: Movies |
       ----------------------------------------------------------
       |           A            B            C          D       |
       |1     The Bodyguard   1992           9                  |
       |2     The Matrix      1999           8                  |
       |3                                                       |
       |4     Skyfall         2012           9                  |
       ----------------------------------------------------------

       PS C:\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Movies -HeaderName 'Movie name', 'Year', 'Rating', 'Genre'

       Movie name: The Bodyguard
       Year      : 1992
       Rating    : 9
       Genre     :

       Movie name: The Matrix
       Year      : 1999
       Rating    : 8
       Genre     :

       Movie name:
       Year      :
       Rating    :
       Genre     :

       Movie name: Skyfall
       Year      : 2012
       Rating    : 9
       Genre     :

       Notice that empty rows are imported and that data for the property 'Genre' is not present in the worksheet. As such, the 'Genre' property will be blanc for all objects.

    .EXAMPLE
       Import data from an Excel worksheet. One object is created for each row. The property names of the objects are automatically generated by using the switch ‘-NoHeader’ (P1, P@, P#, ..). The switch ‘-DataOnly’ will speed up the import because empty rows and empty columns are not imported.

       ----------------------------------------------------------
       | File: Movies.xlsx            -           Sheet: Movies |
       ----------------------------------------------------------
       |           A            B            C          D       |
       |1     The Bodyguard   1992           9                  |
       |2     The Matrix      1999           8                  |
       |3                                                       |
       |4     Skyfall         2012           9                  |
       ----------------------------------------------------------

       PS C:\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Movies –NoHeader -DataOnly

       P1: The Bodyguard
       P2: 1992
       P3: 9

       P1: The Matrix
       P2: 1999
       P3: 8

       P1: Skyfall
       P2: 2012
       P3: 9

       Notice that empty rows and empty columns are not imported.

.EXAMPLE
       Import data from an Excel worksheet. One object is created for each row. The property names are provided with the ‘-HeaderName’ parameter. The import will start from row 2 and empty columns and rows are not imported.

       ----------------------------------------------------------
       | File: Movies.xlsx            -           Sheet: Actors |
       ----------------------------------------------------------
       |           A           B           C            D       |
       |1     Chuck                     Norris       California |
       |2                                                       |
       |3     Jean-Claude               Vandamme     Brussels   |
       ----------------------------------------------------------

       PS C:\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Actors -DataOnly -HeaderName 'FirstName', 'SecondName', 'City' –StartRow 2

       FirstName : Jean-Claude
       SecondName: Vandamme
       City      : Brussels

       Notice that only 1 object is imported with only 3 properties. Column B and row 2 are empty and have been disregarded by using the switch '-DataOnly'. The property names have been named with the values provided with the parameter '-HeaderName'. Row number 1 with ‘Chuck Norris’ has not been imported, because we started the import from row 2 with the parameter ‘-StartRow 2’.

   .LINK
       https://github.com/dfinke/ImportExcel

   .NOTES
  #>

    [CmdLetBinding(DefaultParameterSetName)]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword","")]
    Param (
        [Alias('FullName')]
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline, Position=0, Mandatory)]
        [ValidateScript( {(Test-Path -Path $_ -PathType Leaf) -and ($_ -match '.xls$|.xlsx$|.xlsm$')})]
        [String]$Path,
        [Alias('Sheet')]
        [Parameter(Position=1)]
        [ValidateNotNullOrEmpty()]
        [String]$WorksheetName,
        [Parameter(ParameterSetName='B', Mandatory)]
        [String[]]$HeaderName ,
        [Parameter(ParameterSetName='C', Mandatory)]
        [Switch]$NoHeader     ,
        [Alias('HeaderRow','TopRow')]
        [ValidateRange(1, 9999)]
        [Int]$StartRow = 1,
        [Alias('StopRow','BottomRow')]
        [Int]$EndRow ,
        [Alias('LeftColumn')]
        [Int]$StartColumn = 1,
        [Alias('RightColumn')]
        [Int]$EndColumn  ,
        [Switch]$DataOnly,
        [ValidateNotNullOrEmpty()]
        [String]$Password
    )
    Begin {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        Function Get-PropertyNames {
            <#
            .SYNOPSIS
                Create objects containing the column number and the column name for each of the different header types.
            #>

            Param (
                [Parameter(Mandatory)]
                [Int[]]$Columns,
                [Parameter(Mandatory)]
                [Int]$StartRow
            )

            Try {
                if ($NoHeader) {
                    $i = 0
                    foreach ($C in $Columns) {
                        $i++
                        $C | Select-Object @{N='Column'; E={$_}}, @{N='Value'; E={'P' + $i}}
                    }
                }
                elseif ($HeaderName) {
                    $i = 0
                    foreach ($H in $HeaderName) {
                        $H | Select-Object @{N='Column'; E={$Columns[$i]}}, @{N='Value'; E={$H}}
                        $i++
                    }
                }
                else {
                    if ($StartRow -eq 0) {
                        throw 'The top row can never be equal to 0 when we need to retrieve headers from the worksheet.'
                    }

                    foreach ($C in $Columns) {
                        $Worksheet.Cells[$StartRow,$C] | Where-Object {$_.Value} | Select-Object @{N='Column'; E={$C}}, Value
                    }
                }
            }
            Catch {
                throw "Failed creating property names: $_"
            }
        }
    }

    Process {
        #region Open file
        try {
            $Path = (Resolve-Path $Path).ProviderPath
            Write-Verbose "Import Excel workbook '$Path' with worksheet '$Worksheetname'"
            $Stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Path, 'Open', 'Read', 'ReadWrite'
        } 
        Catch {throw "Could not open $Path ; $_ "} 

        if ($Password) {
            Try   {$Excel = New-Object -TypeName OfficeOpenXml.ExcelPackage 
                   $excel.Load( $Stream,$Password)}
            Catch { throw "Could not read $Path with the provided password." }
        }
        else {
            try   {$Excel = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Stream}
            Catch {throw "Failed to read $Path"}
        }
        #endregion
        Try {
            #region Select worksheet
            if ($WorksheetName) {
                if (-not ($Worksheet = $Excel.Workbook.Worksheets[$WorkSheetName])) {
                    throw "Worksheet '$WorksheetName' not found, the workbook only contains the worksheets '$($Excel.Workbook.Worksheets)'. If you only wish to select the first worksheet, please remove the '-WorksheetName' parameter."
                }
            }
            else {
                $Worksheet = $Excel.Workbook.Worksheets | Select-Object -First 1
            }
            #endregion
            Write-Debug $sw.Elapsed.TotalMilliseconds
            #region Get rows and columns
            #If we are doing dataonly it is quicker to work out which rows to ignore before processing the cells.
            if (-not $EndRow   ) {$EndRow    = $Worksheet.Dimension.End.Row    }
            if (-not $EndColumn) {$EndColumn = $Worksheet.Dimension.End.Column }
            $endAddress = [OfficeOpenXml.ExcelAddress]::TranslateFromR1C1("R[$EndRow]C[$EndColumn]",0,0)
            if ($DataOnly) {
                # If we are supplying headers or letting the module generate them, start at StartRow
                # Else we are using headers from the data so startrow will be the headerrow so examine data from startRow + 1,
                if ($NoHeader -or $HeaderName) {$range = "A" + ($StartRow     ) + ":" + $endAddress }
                else                           {$range = "A" + ($StartRow + 1 ) + ":" + $endAddress }
                #We're going to look at every cell and build 2 hash tables holding rows & columns which contain data.
                #Want to Avoid 'select unique' operations & large Sorts, becuse time time taken increases with square
                #of number of items (PS uses heapsort at large size). Instead keep a list of what we have seen,
                #using Hash tables: "we've seen it" is all we need, no need to worry about "seen it before" / "Seen it many times".
                $colHash = @{}
                $rowHash = @{}
                foreach ($cell in $Worksheet.Cells[$range]) {
                     if ($null -ne $cell.Value ) {$colHash[$cell.Start.Column]=1;  $rowHash[$cell.Start.row]=1 }
                }
                $rows    = (   $StartRow..$EndRow   ).Where({$rowHash[$_]})
                $columns = ($StartColumn..$EndColumn).Where({$colHash[$_]})
            }
            else {
                $Columns = $StartColumn..$EndColumn  ;              if ($StartColumn -gt $EndColumn) {Write-Warning -Message "Selecting columns $StartColumn to $EndColumn might give odd results."}
                if ($NoHeader -or $HeaderName)  {$Rows = (    $StartRow)..$EndRow ; if ($StartRow    -gt $EndRow)    {Write-Warning -Message "Selecting rows $StartRow to $EndRow might give odd results."} }
                else                            {$Rows = (1 + $StartRow)..$EndRow ; if ($StartRow    -ge $EndRow)    {Write-Warning -Message "Selecting $StartRow as the header with data in $(1+$StartRow) to $EndRow might give odd results."}}
            }
            #endregion
            #region Create property names
            if ((-not $Columns) -or (-not ($PropertyNames = Get-PropertyNames -Columns $Columns -StartRow $StartRow))) {
                throw "No column headers found on top row '$StartRow'. If column headers in the worksheet are not a requirement then please use the '-NoHeader' or '-HeaderName' parameter."
            }
            if ($Duplicates = $PropertyNames | Group-Object Value | Where-Object Count -GE 2) {
                throw "Duplicate column headers found on row '$StartRow' in columns '$($Duplicates.Group.Column)'. Column headers must be unique, if this is not a requirement please use the '-NoHeader' or '-HeaderName' parameter."
            }
            #endregion
            Write-Debug $sw.Elapsed.TotalMilliseconds
            if (-not $Rows) {
                Write-Warning "Worksheet '$WorksheetName' in workbook '$Path' contains no data in the rows after top row '$StartRow'"
            }
            else {
                #region Create one object per row
                foreach ($R in $Rows) {
                    Write-Verbose "Import row '$R'"
                    $NewRow = [Ordered]@{}

                    foreach ($P in $PropertyNames) {
                       $NewRow[$P.Value] = $Worksheet.Cells[$R, $P.Column].Value
                       Write-Verbose "Import cell '$($Worksheet.Cells[$R, $P.Column].Address)' with property name '$($p.Value)' and value '$($Worksheet.Cells[$R, $P.Column].Value)'."
                    }

                    [PSCustomObject]$NewRow
                }
                #endregion
            }
            Write-Debug $sw.Elapsed.TotalMilliseconds
        }
        Catch {
            throw "Failed importing the Excel workbook '$Path' with worksheet '$Worksheetname': $_"
        }
        Finally {
            $Stream.Close()
            $Stream.Dispose()
            $Excel.Dispose()
            $Excel = $null
        }
    }
}

function ConvertFrom-ExcelSheet {
    <#
      .Synopsis
        Reads an Excel file an converts the data to a delimited text file.

      .Example
        ConvertFrom-ExcelSheet .\TestSheets.xlsx .\data
        Reads each sheet in TestSheets.xlsx and outputs it to the data directory as the sheet name with the extension .txt.

      .Example
        ConvertFrom-ExcelSheet .\TestSheets.xlsx .\data sheet?0
        Reads and outputs sheets like Sheet10 and Sheet20 form TestSheets.xlsx and outputs it to the data directory as the sheet name with the extension .txt.
    #>

    [CmdletBinding()]
    param
    (
        [Alias("FullName")]
        [Parameter(Mandatory = $true)]
        [String]
        $Path,
        [String]
        $OutputPath = '.\',
        [String]
        $SheetName="*",
        [ValidateSet('ASCII', 'BigEndianUniCode','Default','OEM','UniCode','UTF32','UTF7','UTF8')]
        [string]
        $Encoding = 'UTF8',
        [ValidateSet('.txt', '.log','.csv')]
        [string]
        $Extension = '.csv',
        [ValidateSet(';', ',')]
        [string]
        $Delimiter = ';'
    )

    $Path = (Resolve-Path $Path).Path
    $stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Path,"Open","Read","ReadWrite"
    $xl = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $stream
    $workbook = $xl.Workbook

    $targetSheets = $workbook.Worksheets | Where-Object {$_.Name -like $SheetName}

    $params = @{} + $PSBoundParameters
    $params.Remove("OutputPath")
    $params.Remove("SheetName")
    $params.Remove('Extension')
    $params.NoTypeInformation = $true

    Foreach ($sheet in $targetSheets)
    {
        Write-Verbose "Exporting sheet: $($sheet.Name)"

        $params.Path = "$OutputPath\$($Sheet.Name)$Extension"

        Import-Excel $Path -Sheet $($sheet.Name) | Export-Csv @params
    }

    $stream.Close()
    $stream.Dispose()
    $xl.Dispose()
}

function Export-MultipleExcelSheets {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword","")]
    param(
        [Parameter(Mandatory=$true)]
        $Path,
        [Parameter(Mandatory=$true)]
        [hashtable]$InfoMap,
        [string]$Password,
        [Switch]$Show,
        [Switch]$AutoSize
    )

    $parameters = @{}+$PSBoundParameters
    $parameters.Remove("InfoMap")
    $parameters.Remove("Show")

    $parameters.Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)

    foreach ($entry in $InfoMap.GetEnumerator()) {
        Write-Progress -Activity "Exporting" -Status "$($entry.Key)"
        $parameters.WorkSheetname=$entry.Key

        & $entry.Value | Export-Excel @parameters
    }

    if($Show) {Invoke-Item $Path}
}

Function WorksheetArgumentCompleter {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    $xlPath        = $fakeBoundParameter['Path']
    if (Test-Path -Path $xlPath) {
        $xlpkg = Open-ExcelPackage -Path $xlPath
        $WorksheetNames = $xlPkg.Workbook.Worksheets.Name
        Close-ExcelPackage -nosave -ExcelPackage $xlpkg
        $WorksheetNames.where({$_ -like "*$wordToComplete*"}) | foreach-object {
            New-Object -TypeName System.Management.Automation.CompletionResult -ArgumentList "'$_'",
                            $_ , ([System.Management.Automation.CompletionResultType]::ParameterValue) ,$_
        }
    }
}
If (Get-Command -ErrorAction SilentlyContinue -name Register-ArgumentCompleter) {
    Register-ArgumentCompleter -CommandName 'Import-Excel' -ParameterName 'WorksheetName' -ScriptBlock $Function:WorksheetArgumentCompleter
}
