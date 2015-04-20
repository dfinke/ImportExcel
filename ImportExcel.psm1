Add-Type -Path "$($PSScriptRoot)\EPPlus.dll"

function Import-Excel {
    param(
        [Parameter(ValueFromPipelineByPropertyName)]
        $FullName,
        $Sheet=1,
        [string[]]$Header
    )

    Process {

        $FullName = (Resolve-Path $FullName).Path
        write-debug "target excel file $($FullName)"

        $xl = New-Object OfficeOpenXml.ExcelPackage $FullName

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
                $Name    = $Header[$Column]
                $h.$Name = $worksheet.Cells[$Row,($Column+1)].Text
            }
            [PSCustomObject]$h
        }

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

        $params.Path = "$($OutputPath)\$($Sheet.Name)$($Extension)"

        Import-Excel $Path -Sheet $($sheet.Name) | Export-Csv @params
    }
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
    #>
    param(
        [Parameter(Mandatory)]
        $Path,
        [Parameter(ValueFromPipeline)]
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
        [string[]]$PivotData,
        [string]$Password,
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType="Pie",
        [Switch]$IncludePivotTable,
        [Switch]$IncludePivotChart,
        [Switch]$AutoSize,
        [Switch]$Show,
        [Switch]$NoClobber
    )

    Begin {
        try {
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            $pkg = New-Object OfficeOpenXml.ExcelPackage $Path

            if($pkg.Workbook.Worksheets[$WorkSheetname]) {
                if($NoClobber) {
                    $AlreadyExists = $true
                    throw ""
                } else {
                    $pkg.Workbook.Worksheets.delete($WorkSheetname)
                }
            }

            $ws  = $pkg.Workbook.Worksheets.Add($WorkSheetname)

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
            $ws.Cells[$Row, $ColumnIndex].Value = $TargetData.$Name
            $ColumnIndex += 1
        }
    }

    End {

        if($AutoSize) {$ws.Cells.AutoFitColumns()}

        if($IncludePivotTable) {
            $pivotTableName = $WorkSheetname + "PivotTable"
            $wsPivot = $pkg.Workbook.Worksheets.Add($pivotTableName)
            $wsPivot.View.TabSelected = $true

            $pivotTableDataName=$WorkSheetname + "PivotTableData"

            $startAddress=$ws.Dimension.Start.Address
            if($Title) {$startAddress="A2"}

            $range="{0}:{1}" -f $startAddress, $ws.Dimension.End.Address
            $pivotTable = $wsPivot.PivotTables.Add($wsPivot.Cells["A1"], $ws.Cells[$range], $pivotTableDataName)

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
                foreach ($Item in $PivotData) {
                    $null=$pivotTable.DataFields.Add($pivotTable.Fields[$Item])
                }
            }

            if($IncludePivotChart) {
                $chart = $wsPivot.Drawings.AddChart("PivotChart", $ChartType, $pivotTable)
                $chart.SetPosition(1, 0, 6, 0)
                $chart.SetSize(600, 400)
            }
        }

        if($Password) { $ws.Protection.SetPassword($Password) }

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
    $xl = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path
    $workbook = $xl.Workbook

    $targetSheets = $workbook.Worksheets | Where {$_.Name -like $SheetName}

    $params = @{} + $PSBoundParameters
    $params.Remove("OutputPath")
    $params.Remove("SheetName")
    $params.NoTypeInformation = $true

    Foreach ($sheet in $targetSheets)
    {
        Write-Verbose "Exporting sheet: $($sheet.Name)"

        $params.Path = "$($OutputPath)\$($Sheet.Name)$($Extension)"

        Import-Excel $Path -Sheet $($sheet.Name) | Export-Csv @params
    }
}