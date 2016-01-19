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
        [Switch]$NoLegend,
        [Switch]$ShowCategory,
        [Switch]$ShowPercent,
        [Switch]$AutoSize,
        [Switch]$Show,
        [Switch]$NoClobber,
        [Switch]$FreezeTopRow,
        [Switch]$AutoFilter,
        [Switch]$BoldTopRow,
        [Switch]$NoHeader,
        [string]$RangeName,
        [string]$TableName,
        [OfficeOpenXml.Table.TableStyles]$TableStyle="Medium6",
        [Object[]]$ConditionalFormat,
        [Object[]]$ConditionalText,        
        [Object[]]$ExcelChartDefinition,
        [string[]]$HideSheet,
        [Switch]$KillExcel,
        [Switch]$AutoNameRange,
        $StartRow=1,
        $StartColumn=1
    )

    Begin {
        if($KillExcel) {
            Get-Process excel -ErrorAction Ignore | Stop-Process
            while (Get-Process excel -ErrorAction Ignore) {}
        }

        try {
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            if (Test-Path $path) {
                Write-Debug "File `"$Path`" already exists"
            }
            $pkg = New-Object OfficeOpenXml.ExcelPackage $Path

            $ws  = $pkg | Add-WorkSheet -WorkSheetname $WorkSheetname -NoClobber:$NoClobber

            foreach($format in $ConditionalFormat ) {
                $target = "Add$($format.Formatter)"
                $rule = ($ws.ConditionalFormatting).$target($format.Address, $format.IconType)
                $rule.Reverse = $format.Reverse
            }

            # Force at least one cell value
            #$ws.Cells[1, 1].Value = ""


            $Row = $StartRow
            if($Title) {
                $ws.Cells[$Row, $StartColumn].Value = $Title

                $ws.Cells[$Row, $StartColumn].Style.Font.Size = $TitleSize
                $ws.Cells[$Row, $StartColumn].Style.Font.Bold = $TitleBold
                $ws.Cells[$Row, $StartColumn].Style.Fill.PatternType = $TitleFillPattern
                if($TitleBackgroundColor) {
                    $ws.Cells[$Row, $StartColumn].Style.Fill.BackgroundColor.SetColor($TitleBackgroundColor)
                }

                $Row += 1
            }

        } Catch {
            if($AlreadyExists) {
                throw "$WorkSheetname already exists."
            } else {
                throw $Error[0].Exception.Message
            }
        }

        $firstTimeThru = $true
        $isDataTypeValueType=$false
        $pattern = "string|bool|byte|char|decimal|double|float|int|long|sbyte|short|uint|ulong|ushort"
    }

    Process {
        if($firstTimeThru) {
            $firstTimeThru=$false
            $isDataTypeValueType = $TargetData.GetType().name -match "string|bool|byte|char|decimal|double|float|int|long|sbyte|short|uint|ulong|ushort"
        }

        if($isDataTypeValueType) {
            $ColumnIndex = $StartColumn

            $targetCell = $ws.Cells[$Row, $ColumnIndex]

            $r=$null
            $cellValue=$TargetData
            if([double]::tryparse($cellValue, [ref]$r)) {
                $targetCell.Value = $r
            } else {
                $targetCell.Value = $cellValue
            }

            switch ($TargetData.$Name) {
                {$_ -is [datetime]} {$targetCell.Style.Numberformat.Format = "m/d/yy h:mm"}
            }

            $ColumnIndex += 1
            $Row += 1

        } else {
            if(!$Header) {

                $ColumnIndex = $StartColumn

                $Header = $TargetData.psobject.properties.name

                if($NoHeader) {
                    # Don't push the headers to the spread sheet
                    $Row -= 1
                } else {
                    foreach ($Name in $Header) {
                        $ws.Cells[$Row, $ColumnIndex].Value = $name
                        $ColumnIndex += 1
                    }
                }
            }

            $Row += 1
            $ColumnIndex = $StartColumn

            foreach ($Name in $Header) {

                $targetCell = $ws.Cells[$Row, $ColumnIndex]

                $cellValue=$TargetData.$Name

                if($cellValue -is [string] -and $cellValue.StartsWith('=')) {
                    $targetCell.Formula = $cellValue
                } else {

                    $r=$null
                    if([double]::tryparse($cellValue, [ref]$r)) {
                        $targetCell.Value = $r
                    } else {
                        $targetCell.Value = $cellValue
                    }
                }

                switch ($TargetData.$Name) {
                    {$_ -is [datetime]} {$targetCell.Style.Numberformat.Format = "m/d/yy h:mm"}
                }

                #[ref]$uriResult=$null
                #if ([uri]::TryCreate($cellValue, [System.UriKind]::Absolute, $uriResult)) {

                #    $targetCell.Hyperlink = [uri]$cellValue

                #    $namedStyle=$ws.Workbook.Styles.NamedStyles | where {$_.Name -eq 'HyperLink'}
                #    if(!$namedStyle) {
                #        $namedStyle=$ws.Workbook.Styles.CreateNamedStyle("HyperLink")
                #        $namedStyle.Style.Font.UnderLine = $true
                #        $namedStyle.Style.Font.Color.SetColor("Blue")
                #    }

                #    $targetCell.StyleName = "HyperLink"
                #}

                $ColumnIndex += 1
            }
        }
    }

    End {
        if($AutoNameRange) {
            $totalRows=$ws.Dimension.Rows
            $totalColumns=$ws.Dimension.Columns

            foreach($c in 0..($totalColumns-1)) {
                $targetRangeName = "$($Header[$c])"
                $targetColumn = $c+1
                $theCell = $ws.Cells[2,$targetColumn,$totalRows,$targetColumn ]
                $ws.Names.Add($targetRangeName, $theCell) | Out-Null
            }
        }

        $startAddress=$ws.Dimension.Start.Address
        $dataRange="{0}:{1}" -f $startAddress, $ws.Dimension.End.Address

        Write-Debug "Data Range $dataRange"

        if (-not [string]::IsNullOrEmpty($RangeName)) {
            $ws.Names.Add($RangeName, $ws.Cells[$dataRange]) | Out-Null
        }

        if (-not [string]::IsNullOrEmpty($TableName)) {
            #$ws.Tables.Add($ws.Cells[$dataRange], $TableName) | Out-Null
            #"$($StartRow),$($StartColumn),$($ws.Dimension.End.Row-$StartRow),$($Header.Count)"

            $csr=$StartRow
            $csc=$StartColumn
            $cer=$ws.Dimension.End.Row #-$StartRow+1
            $cec=$Header.Count

            $targetRange=$ws.Cells[$csr, $csc, $cer,$cec]

            $tbl = $ws.Tables.Add($targetRange, $TableName)

            $tbl.TableStyle=$TableStyle

            $idx
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

                $chart.DataLabel.ShowCategory=$ShowCategory
                $chart.DataLabel.ShowPercent=$ShowPercent

                if($NoLegend) { $chart.Legend.Remove() }

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

        foreach($Sheet in $HideSheet) {
            $pkg.Workbook.WorkSheets[$Sheet].Hidden="Hidden"

        }

        $chartCount=0
        foreach ($chartDef in $ExcelChartDefinition) {
            $ChartName = "Chart"+(Split-Path -Leaf ([System.IO.path]::GetTempFileName())) -replace 'tmp|\.',''
            $chart = $ws.Drawings.AddChart($ChartName, $chartDef.ChartType)
            $chart.Title.Text = $chartDef.Title

            if($chartDef.NoLegend) {
                $chart.Legend.Remove()
            }
            #$chart.Datalabel.ShowLegendKey = $true
            $chart.Datalabel.ShowCategory  = $chartDef.ShowCategory
            $chart.Datalabel.ShowPercent   = $chartDef.ShowPercent

            $chart.SetPosition($chartDef.Row, $chartDef.RowOffsetPixels,$chartDef.Column, $chartDef.ColumnOffsetPixels)
            $chart.SetSize($chartDef.Width, $chartDef.Height)

            $chartDefCount = @($chartDef.XRange).Count
            if($chartDefCount -eq 1) {
                $Series=$chart.Series.Add($chartDef.YRange, $chartDef.XRange)
                $Series.Header = $chartDef.Header
            } else {
                for($idx=0; $idx -lt $chartDefCount; $idx+=1) {
                    $Series=$chart.Series.Add($chartDef.YRange[$idx], $chartDef.XRange)
                    $Series.Header = $chartDef.Header[$idx]
                }
            }
        }

        if($ConditionalText) {       
            foreach ($targetConditionalText in $ConditionalText) {
                $target = "Add$($targetConditionalText.ConditionalType)"                
                
                $Range=$targetConditionalText.Range
                if(!$Range) { $Range=$ws.Dimension.Address }

                $rule=($ws.Cells[$Range].ConditionalFormatting).$target()
                
                if($targetConditionalText.Text) {
                    $rule.Text = $targetConditionalText.Text
                }                
                
                $rule.Style.Font.Color.Color = $targetConditionalText.ConditionalTextColor
                $rule.Style.Fill.PatternType=$targetConditionalText.PatternType
                $rule.Style.Fill.BackgroundColor.Color=$targetConditionalText.BackgroundColor
           }
        }
        
        $pkg.Save()
        $pkg.Dispose()

        if($Show) {Invoke-Item $Path}
    }
}
