Add-Type -Path "$($PSScriptRoot)\EPPlus.dll"

. $PSScriptRoot\Export-Excel.ps1
. $PSScriptRoot\New-ConditionalFormattingIconSet.ps1

function Import-Excel {
    param(
		[Alias("FullName")]
        [Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, Mandatory)]
        $Path,
        $Sheet=1,
        [string[]]$Header
    )

    Process {

        $Path = (Resolve-Path $Path).Path
        write-debug "target excel file $Path"
		
		$stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Path,"Open","Read","ReadWrite"
        $xl = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $stream

        $workbook  = $xl.Workbook

        $worksheet=$workbook.Worksheets[$Sheet]
        $dimension=$worksheet.Dimension

        $Rows=$dimension.Rows
        $Columns=$dimension.Columns

        if(!$Header) {
            $Header = foreach ($Column in 1..$Columns) {
                $worksheet.Cells[1,$Column].Text
            }
        }

        foreach ($Row in 2..$Rows) {
            $h=[Ordered]@{}
            foreach ($Column in 0..($Columns-1)) {
                if($Header[$Column].Length -gt 0) {
                    $Name    = $Header[$Column]
                    $h.$Name = $worksheet.Cells[$Row,($Column+1)].Value
                }
            }
            [PSCustomObject]$h
        }
		
		$stream.Close()
		$stream.Dispose()
        $xl.Dispose()
        $xl = $null
    }
}

function Export-ExcelSheet {

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [String]
        $Path,
        [String]
        $OutputPath = '.\',
        [String]
        $SheetName,
        [string]
        $Encoding = 'UTF8',
        [string]
        $Extension = '.txt',
        [string]
        $Delimiter = ';'
    )

    $Path = (Resolve-Path $Path).Path
    $xl = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path
    $workbook = $xl.Workbook

    $targetSheets = $workbook.Worksheets | Where {$_.Name -Match $SheetName}

    $params = @{} + $PSBoundParameters
    $params.Remove("OutputPath")
    $params.Remove("SheetName")
    $params.NoTypeInformation = $true

    Foreach ($sheet in $targetSheets)
    {
        Write-Verbose "Exporting sheet: $($sheet.Name)"

        $params.Path = "$OutputPath\$($Sheet.Name)$Extension"

        Import-Excel $Path -Sheet $($sheet.Name) | Export-Csv @params -Encoding $Encoding
    }

    $xl.Dispose()
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
    if($ExcelPackage.Workbook.Worksheets[$WorkSheetname]) {
        if($NoClobber) {
            $AlreadyExists = $true
            Write-Error "Worksheet `"$WorkSheetname`" already exists."
        } else {
            Write-Debug "Worksheet `"$WorkSheetname`" already exists. Deleting"
            $ExcelPackage.Workbook.Worksheets.Delete($WorkSheetname)
        }
    }

    $ExcelPackage.Workbook.Worksheets.Add($WorkSheetname)
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
        [string]
        $Encoding = 'UTF8',
        [string]
        $Extension = '.txt',
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
    $params.NoTypeInformation = $true

    Foreach ($sheet in $targetSheets)
    {
        Write-Verbose "Exporting sheet: $($sheet.Name)"

        $params.Path = "$OutputPath\$($Sheet.Name)$Extension"

        Import-Excel $Path -Sheet $($sheet.Name) | Export-Csv @params -Encoding $Encoding
    }
	
	$stream.Close()
	$stream.Dispose()
    $xl.Dispose()
}

function Export-MultipleExcelSheets {
    param(
        [Parameter(Mandatory)]
        $Path,
        [Parameter(Mandatory)]
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