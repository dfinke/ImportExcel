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
        [string[]]$PivotRows,
        [string[]]$PivotColumns,
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
        [string]$TableName,
        [Object[]]$ConditionalFormat
    )

    Begin {
        try {
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            if (Test-Path $path) {
                Write-Debug "File `"$Path`" already exists"
            }
            $pkg = New-Object OfficeOpenXml.ExcelPackage $Path

            $ws  = $pkg | Add-WorkSheet -WorkSheetname $WorkSheetname -NoClobber:$NoClobber

            foreach($format in $ConditionalFormat ) {
                #$obj = [PSCustomObject]@{
                #    Address   = $Address
                #    Formatter = $ConditionalFormat
                #    IconType  = $bp.IconType
                #}

                $target = "Add$($format.Formatter)"
                $rule = ($ws.ConditionalFormatting).$target($format.Address, $format.IconType)
                $rule.Reverse = $format.Reverse
            }

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
