function Import-ByColumns    {
<#
    .synopsis
        Works like Import-Excel but with data in columns instead of the conventional rows.
    .Description.
        Import-excel will read the sample file in this folder like this
        > Import-excel  FruitCity.xlsx | ft *
            GroupAs Apple Orange Banana
            ------- ----- ------ ------
            London      1      4      9
            Paris       2      4     10
            NewYork     6      5     11
            Munich      7      8     12
        Import-ByColumns transposes it
        > Import-Bycolumns FruitCity.xlsx | ft *
            GroupAs London Paris NewYork Munich
            ------- ------ ----- ------- ------
            Apple   1      2     6       7
            Orange  4      4     5       8
            Banana  9      10    11      12
    .Example
        C:\> Import-Bycolumns -path .\VM_Build_Example.xlsx -StartRow 7 -EndRow 21 -EndColumn  7  -HeaderName Desc,size,type,
            cpu,ram,NetAcc,OS,OSDiskSize,DataDiskSize,LogDiskSize,TempDbDiskSize,BackupDiskSize,ImageDiskDize,AzureBackup,AzureReplication | ft -a *

        This reads a spreadsheet which has a block from row 7 to 21 containing 14 properties of virtual machines.
        The properties names are in column A and the 6 VMS are in columns B-G
        Because the property names are written for easy reading by the person completing the spreadsheet, they are replaced with new names.
        All the parameters work as they would for Import-Excel
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "")]
    param(
        [Alias('FullName')]
        [Parameter(ParameterSetName = "PathA", Mandatory, ValueFromPipelineByPropertyName, ValueFromPipeline, Position = 0 )]
        [Parameter(ParameterSetName = "PathB", Mandatory, ValueFromPipelineByPropertyName, ValueFromPipeline, Position = 0 )]
        [Parameter(ParameterSetName = "PathC", Mandatory, ValueFromPipelineByPropertyName, ValueFromPipeline, Position = 0 )]
        [String]$Path,

        [Parameter(ParameterSetName = "PackageA", Mandatory)]
        [Parameter(ParameterSetName = "PackageB", Mandatory)]
        [Parameter(ParameterSetName = "PackageC", Mandatory)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,

        [Alias('Sheet')]
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [String]$WorksheetName,

        [Parameter(ParameterSetName = 'PathB'   , Mandatory)]
        [Parameter(ParameterSetName = 'PackageB', Mandatory)]
        [String[]]$HeaderName ,
        [Parameter(ParameterSetName = 'PathC'   , Mandatory)]
        [Parameter(ParameterSetName = 'PackageC', Mandatory)]
        [Switch]$NoHeader,

        [Alias('TopRow')]
        [ValidateRange(1, 9999)]
        [Int]$StartRow = 1,

        [Alias('StopRow', 'BottomRow')]
        [Int]$EndRow ,

        [Alias('LeftColumn','LabelColumn')]
        [Int]$StartColumn = 1,

        [Int]$EndColumn,
        [switch]$DataOnly,
        [switch]$AsHash,

        [ValidateNotNullOrEmpty()]
        [String]$Password
    )
    function Get-PropertyNames {
         <#
            .SYNOPSIS
                Create objects containing the row number and the row name for each of the different header types.
        #>
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "Name would be incorrect, and command is not exported")]
        param(
            [Parameter(Mandatory)]
            [Int[]]$Rows,
            [Parameter(Mandatory)]
            [Int]$StartColumn
        )
        if ($HeaderName) {
            $i = 0
            foreach ($h in $HeaderName) {
                $h | Select-Object @{n='Row'; e={$rows[$i]}}, @{n='Value'; e={$h} }
                $i++
            }
        }
        elseif ($NoHeader) {
            $i = 0
            foreach ($r in $rows) {
                $i++
                $r | Select-Object @{n='Row'; e={$_}}, @{n='Value'; e={"P$i"} }
            }
        }
        else {
            foreach ($r in $Rows) {
                #allow "False" or "0" to be  headings
                 $Worksheet.Cells[$r, $StartColumn] | Where-Object {-not [string]::IsNullOrEmpty($_.Value) } | Select-Object @{n='Row'; e={$r} }, Value
            }
        }
    }

#region open file if necessary, find worksheet and ensure we have start/end row/columns
    if   ($Path -and -not $ExcelPackage -and $Password) {
        $ExcelPackage = Open-ExcelPackage -Path $Path -Password $Password
    }
    elseif ($Path -and -not $ExcelPackage ) {
        $ExcelPackage = Open-ExcelPackage -Path $Path
    }
    if (-not $ExcelPackage) {
        throw 'Could not get an Excel workbook to work on' ; return
    }

    if     (-not  $WorksheetName) { $Worksheet = $ExcelPackage.Workbook.Worksheets[1] }
    elseif (-not ($Worksheet = $ExcelPackage.Workbook.Worksheets[$WorkSheetName])) {
        throw "Worksheet '$WorksheetName' not found, the workbook only contains the worksheets '$($ExcelPackage.Workbook.Worksheets)'. If you only wish to select the first worksheet, please remove the '-WorksheetName' parameter." ; return
    }

    if (-not $EndRow   ) { $EndRow    = $Worksheet.Dimension.End.Row }
    if (-not $EndColumn) { $EndColumn = $Worksheet.Dimension.End.Column }
#endregion

    $Rows    = $Startrow .. $EndRow  ;
    $Columns = (1 + $StartColumn)..$EndColumn

    if ((-not $rows) -or (-not ($PropertyNames = Get-PropertyNames -Rows $Rows -StartColumn $StartColumn))) {
        throw "No headers found in left coulmn '$Startcolumn'. "; return
    }
    if (-not $Columns) {
        Write-Warning "Worksheet '$WorksheetName' in workbook contains no data in the rows after left column '$StartColumn'"
    }
    else {
        foreach ($c in $Columns) {
            $NewColumn = [Ordered]@{ }
            foreach ($p in $PropertyNames) {
                $NewColumn[$p.Value] = $Worksheet.Cells[$p.row,$c].text
            }
            if     ($AsHash)                                        {$NewColumn}
            elseif (($NewColumn.Values -ne "") -or -not $dataonly)  {[PSCustomObject]$NewColumn}
        }
    }
}
