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

function Export-Excel {
    param(
        [Parameter(Mandatory)]
        $Path,
        [Parameter(ValueFromPipeline)]
        $TargetData,
        $WorkSheetname="Sheet1",
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
        } Catch {
            if($AlreadyExists) {
                throw "$WorkSheetname already exists."
            } else {
                throw $Error[0].Exception.InnerException
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
            #$wsPivot.View.TabSelected = $true

            $pivotTableDataName=$WorkSheetname + "PivotTableData"
            $range="{0}:{1}" -f $ws.Dimension.Start.Address, $ws.Dimension.End.Address
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

        $ws.Protection.SetPassword($Password)

        $pkg.Save()
        $pkg.Dispose()

        if($Show) {Invoke-Item $Path}
    }
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
