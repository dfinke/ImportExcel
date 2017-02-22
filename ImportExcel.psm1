Add-Type -Path "$($PSScriptRoot)\EPPlus.dll"

. $PSScriptRoot\Export-Excel.ps1
. $PSScriptRoot\New-ConditionalFormattingIconSet.ps1
. $PSScriptRoot\New-ConditionalText.ps1
. $PSScriptRoot\Export-ExcelSheet.ps1
. $PSScriptRoot\New-ExcelChart.ps1
. $PSScriptRoot\Invoke-Sum.ps1
. $PSScriptRoot\InferData.ps1
. $PSScriptRoot\Get-ExcelColumnName.ps1
. $PSScriptRoot\Get-XYRange.ps1
. $PSScriptRoot\Charting.ps1
. $PSScriptRoot\New-PSItem.ps1
. $PSScriptRoot\Pivot.ps1
. $PSScriptRoot\Get-ExcelSheetInfo.ps1
. $PSScriptRoot\Get-ExcelWorkbookInfo.ps1
. $PSScriptRoot\Get-HtmlTable.ps1
. $PSScriptRoot\Import-Html.ps1
. $PSScriptRoot\Get-Range.ps1
. $PSScriptRoot\TrackingUtils.ps1
. $PSScriptRoot\Copy-ExcelWorkSheet.ps1
. $PSScriptRoot\Set-CellStyle.ps1

if($PSVersionTable.PSVersion.Major -ge 5) {
    . $PSScriptRoot\plot.ps1

    function New-Plot {
        [OutputType([PSPlot])]
        param()

        [psplot]::new()
    }

} else {
    Write-Warning "PowerShell 5 is required for plot.ps1"
    Write-Warning "PowerShell Excel is ready, except for that functionality"
}


function Import-Excel {
    <#
    .SYNOPSIS
        Read the content of an Excel sheet.
 
    .DESCRIPTION 
        The Import-Excel cmdlet reads the content of an Excel worksheet and creates one object for each row. This is done without using Microsoft Excel in the background but by using the .NET EPPLus.dll. You can also automate the creation of Pivot Tables and Charts.
 
    .PARAMETER Path 
        Specifies the path to the Excel file.
 
    .PARAMETER WorkSheetname
        Specifies the name of the worksheet in the Excel workbook. 
        
    .PARAMETER HeaderRow
        Specifies custom header names for columns.

    .PARAMETER Header
        Specifies the title used in the worksheet. The title is placed on the first line of the worksheet.

    .PARAMETER NoHeader
        When used we generate our own headers (P1, P2, P3, ..) instead of the ones defined in the first row of the Excel worksheet.

    .PARAMETER DataOnly
        When used we will only generate objects for rows that contain text values, not for empty rows or columns.
 
    .EXAMPLE
        Import-Excel -WorkSheetname 'Statistics' -Path 'E:\Finance\Company results.xlsx'
        Imports all the information found in the worksheet 'Statistics' of the Excel file 'Company results.xlsx'

    .LINK
        https://github.com/dfinke/ImportExcel
    #>
    param(
        [Alias("FullName")]
        [Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, Mandatory=$true)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        $Path,
        [Alias("Sheet")]
        $WorkSheetname=1,
        [int]$HeaderRow=1,
        [string[]]$Header,
        [switch]$NoHeader,
        [switch]$DataOnly
    )

    Process {

        $Path = (Resolve-Path $Path).ProviderPath
        write-debug "target excel file $Path"

        $stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Path,"Open","Read","ReadWrite"
        $xl = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $stream

        $workbook  = $xl.Workbook

        $worksheet=$workbook.Worksheets[$WorkSheetname]
        $dimension=$worksheet.Dimension

        $Rows=$dimension.Rows
        $Columns=$dimension.Columns

        if ($NoHeader) {
            if ($DataOnly) {
                $CellsWithValues = $worksheet.Cells | where Value

                $Script:i = 0
                $ColumnReference = $CellsWithValues | Select-Object -ExpandProperty End | Group-Object Column |
                    Select-Object @{L='Column';E={$_.Name}}, @{L='NewColumn';E={$Script:i++; $Script:i}}
                
                $CellsWithValues | Select-Object -ExpandProperty End | Group-Object Row | ForEach-Object {    
                    $newRow = [Ordered]@{}
                    
                    foreach ($C in $ColumnReference) {
                        $newRow."P$($C.NewColumn)" = $worksheet.Cells[($_.Name),($C.Column)].Value
                    }

                    [PSCustomObject]$newRow
                }
            }
            else {
                foreach ($Row in 0..($Rows-1)) {
                    $newRow = [Ordered]@{}
                    foreach ($Column in 0..($Columns-1)) {
                        $propertyName = "P$($Column+1)"
                        $newRow.$propertyName = $worksheet.Cells[($Row+1),($Column+1)].Value
                    }

                    [PSCustomObject]$newRow
                }
            }
        } 
        else {
            if (!$Header) {
                $Header = foreach ($Column in 1..$Columns) {
                    $worksheet.Cells[$HeaderRow,$Column].Value
                }
            }

            if ($Rows -eq 1) {
                $Header | ForEach {$h=[Ordered]@{}} {$h.$_=''} {[PSCustomObject]$h}
            } 
            else {
                if ($DataOnly) {
                    $CellsWithValues = $worksheet.Cells | where {$_.Value -and ($_.End.Row -ne 1)}

                    $Script:i = -1
                    $ColumnReference = $CellsWithValues | Select-Object -ExpandProperty End | Group-Object Column |
                        Select-Object @{L='Column';E={$_.Name}}, @{L='NewColumn';E={$Script:i++; $Header[$Script:i]}}
                
                    $CellsWithValues | Select-Object -ExpandProperty End | Group-Object Row | ForEach-Object {    
                        $newRow = [Ordered]@{}
                    
                        foreach ($C in $ColumnReference) {
                            $newRow."$($C.NewColumn)" = $worksheet.Cells[($_.Name),($C.Column)].Value
                        }

                        [PSCustomObject]$newRow
                    }
                }
                else {
                    foreach ($Row in ($HeaderRow+1)..$Rows) {
                        $h=[Ordered]@{}
                                            foreach ($Column in 0..($Columns-1)) {
                        if($Header[$Column].Length -gt 0) {
                            $Name    = $Header[$Column]
                            $h.$Name = $worksheet.Cells[$Row,($Column+1)].Value
                        }
                    }
                        [PSCustomObject]$h
                    }
                }
            }
        }

        $stream.Close()
        $stream.Dispose()
        $xl.Dispose()
        $xl = $null
    }
}

function Add-WorkSheet {
    param(
        #TODO Use parametersets to allow a workbook to be passed instead of a package
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [OfficeOpenXml.ExcelPackage] $ExcelPackage,
        [Parameter(Mandatory=$true)]
        [string] $WorkSheetname,
        [Switch] $NoClobber
    )

    $ws = $ExcelPackage.Workbook.Worksheets[$WorkSheetname]

    if(!$ws) {
        Write-Verbose "Add worksheet '$WorkSheetname'"
        $ws=$ExcelPackage.Workbook.Worksheets.Add($WorkSheetname)
    }

    return $ws
}

function ConvertFrom-ExcelSheet {
    <#
        .Synopsis
        Reads an Excel file an converts the data to a delimited text file

        .Example
        ConvertFrom-ExcelSheet .\TestSheets.xlsx .\data
        Reads each sheet in TestSheets.xlsx and outputs it to the data directory as the sheet name with the extension .txt

        .Example
        ConvertFrom-ExcelSheet .\TestSheets.xlsx .\data sheet?0
        Reads and outputs sheets like Sheet10 and Sheet20 form TestSheets.xlsx and outputs it to the data directory as the sheet name with the extension .txt
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

    $targetSheets = $workbook.Worksheets | Where {$_.Name -like $SheetName}

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
