Add-Type -Path "$($PSScriptRoot)\EPPlus.dll"

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

function Export-Excel {
    <#
        .Synopsis
        .Example
        gsv | Export-Excel .\test.xlsx
        .Example
        ps | Export-Excel .\test.xlsx -show\
        .Example
        ps | Export-Excel .\test.xlsx -WorkSheetname Processes -IncludePivotTable -Show -PivotRows Company -PivotData PM
        .Example
        ps | Export-Excel .\test.xlsx -WorkSheetname Processes -ChartType PieExploded3D -IncludePivotChart -IncludePivotTable -Show -PivotRows Company -PivotData PM
        .Example
        Remove-Item "c:\temp\test.xlsx" -ErrorAction Ignore
        Get-Service | Export-Excel "c:\temp\test.xlsx"  -Show -IncludePivotTable -PivotRows status -PivotData @{status='count'}
    #>
    param(
        [Parameter(Mandatory=$true)]
        $Path,
        [Parameter(ValueFromPipeline=$true)]
        $TargetData,
        [string]$WorkSheetname="Sheet1",
        [string]$Title,
        [OfficeOpenXml.Style.ExcelFillStyle]$TitleFillPattern="None",
        [bool]$TitleBold,
        [int]$TitleSize=22,
        [System.Drawing.Color]$TitleBackgroundColor,
        #[string]$TitleBackgroundColor,
        [string[]]$PivotRows,
        [string[]]$PivotColumns,
        #[string[]]$PivotData,
        $PivotData,
        [string]$Password,
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType="Pie",
        [Switch]$IncludePivotTable,
        [Switch]$IncludePivotChart,
        [Switch]$AutoSize,
        [Switch]$Show,
        [Switch]$NoClobber,
        [Switch]$FreezeTopRow,
        [Switch]$AutoFilter,
        [Switch]$BoldTopRow,
        [string]$RangeName,
        [string]$TableName
    )

    Begin {
        try {
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            if (Test-Path $path) {
                Write-Debug "File `"$Path`" already exists"
            }
            $pkg = New-Object OfficeOpenXml.ExcelPackage $Path

            $ws  = $pkg | Add-WorkSheet -WorkSheetname $WorkSheetname -NoClobber:$NoClobber

            $Row = 1
            if($Title) {
                $ws.Cells[$Row, 1].Value = $Title

                $ws.Cells[$Row, 1].Style.Font.Size = $TitleSize
                $ws.Cells[$Row, 1].Style.Font.Bold = $TitleBold
                $ws.Cells[$Row, 1].Style.Fill.PatternType = $TitleFillPattern
                if($TitleBackgroundColor) {
                    $ws.Cells[$Row, 1].Style.Fill.BackgroundColor.SetColor($TitleBackgroundColor)
                }

                $Row = 2
            }

        } Catch {
            if($AlreadyExists) {
                throw "$WorkSheetname already exists."
            } else {
                throw $Error[0].Exception.Message
            }
        }
    }

    Process {

        if(!$Header) {

            $ColumnIndex = 1
            $Header = $TargetData.psobject.properties.name

            foreach ($Name in $Header) {
                $ws.Cells[$Row, $ColumnIndex].Value = $name
                $ColumnIndex += 1
            }
        }

        $Row += 1
        $ColumnIndex = 1

        foreach ($Name in $Header) {

            $targetCell = $ws.Cells[$Row, $ColumnIndex]

            $cellValue=$TargetData.$Name

            $r=$null
            if([double]::tryparse($cellValue, [ref]$r)) {
                $targetCell.Value = $r
            } else {
                $targetCell.Value = $cellValue
            }

            switch ($TargetData.$Name) {
                {$_ -is [datetime]} {$targetCell.Style.Numberformat.Format = "m/d/yy h:mm"}
            }

            $ColumnIndex += 1
        }
    }

    End {
        $startAddress=$ws.Dimension.Start.Address
        $dataRange="{0}:{1}" -f $startAddress, $ws.Dimension.End.Address
        Write-Debug "Data Range $dataRange"

        if (-not [string]::IsNullOrEmpty($RangeName)) {
            $ws.Names.Add($RangeName, $ws.Cells[$dataRange]) | Out-Null
        }
        if (-not [string]::IsNullOrEmpty($TableName)) {
            $ws.Tables.Add($ws.Cells[$dataRange], $TableName) | Out-Null
        }

        if($IncludePivotTable) {
            $pivotTableName = $WorkSheetname + "PivotTable"
            $wsPivot = $pkg | Add-WorkSheet -WorkSheetname $pivotTableName -NoClobber:$NoClobber

            $wsPivot.View.TabSelected = $true

            $pivotTableDataName=$WorkSheetname + "PivotTableData"

            if($Title) {$startAddress="A2"}
            $pivotTable = $wsPivot.PivotTables.Add($wsPivot.Cells["A1"], $ws.Cells[$dataRange], $pivotTableDataName)

            if($PivotRows) {
                foreach ($Row in $PivotRows) {
                    $null=$pivotTable.RowFields.Add($pivotTable.Fields[$Row])
                }
            }

            if($PivotColumns) {
                foreach ($Column in $PivotColumns) {
                    $null=$pivotTable.ColumnFields.Add($pivotTable.Fields[$Column])
                }
            }

            if($PivotData) {
                if($PivotData -is [hashtable]) {
                    $PivotData.Keys | % {
                        $df=$pivotTable.DataFields.Add($pivotTable.Fields[$_])
                        $df.Function = $PivotData.$_
                    }
                } else {
                    foreach ($Item in $PivotData) {
                        $df=$pivotTable.DataFields.Add($pivotTable.Fields[$Item])
                        $df.Function = 'Count'
                    }
                }
            }

            if($IncludePivotChart) {
                $chart = $wsPivot.Drawings.AddChart("PivotChart", $ChartType, $pivotTable)
                $chart.SetPosition(1, 0, 6, 0)
                $chart.SetSize(600, 400)
            }
        }

        if($Password) { $ws.Protection.SetPassword($Password) }

        if($AutoFilter) {
            $ws.Cells[$dataRange].AutoFilter=$true
        }

        if($FreezeTopRow) {
            $ws.View.FreezePanes(2,1)
        }

        if($BoldTopRow) {
            $range=$ws.Dimension.Address -replace $ws.Dimension.Rows, "1"
            $ws.Cells[$range].Style.Font.Bold=$true
        }

        if($AutoSize) { $ws.Cells.AutoFitColumns() }

        #$pkg.Workbook.View.ActiveTab = $ws.SheetID

        $pkg.Save()
        $pkg.Dispose()

        if($Show) {Invoke-Item $Path}
    }
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