Function Export-Excel {
    <#
        .SYNOPSIS
            Export data to an Excel work sheet.

        .EXAMPLE
            Get-Service | Export-Excel .\test.xlsx

        .EXAMPLE
            Get-Process | Export-Excel .\test.xlsx -show\

        .EXAMPLE
            Get-Process | Export-Excel .\test.xlsx -WorkSheetname Processes -IncludePivotTable -Show -PivotRows Company -PivotData PM

        .EXAMPLE
            Get-Process | Export-Excel .\test.xlsx -WorkSheetname Processes -ChartType PieExploded3D -IncludePivotChart -IncludePivotTable -Show -PivotRows Company -PivotData PM

        .EXAMPLE
            Remove-Item "c:\temp\test.xlsx" -ErrorAction Ignore
            Get-Service | Export-Excel "c:\temp\test.xlsx"  -Show -IncludePivotTable -PivotRows status -PivotData @{status='count'}
    #>

    [CmdLetBinding()]
    Param(
        #[Parameter(Mandatory=$true)]
        $Path,
        [Parameter(ValueFromPipeline=$true)]
        $TargetData,
        [String]$WorkSheetname='Sheet1',
        [String]$Title,
        [OfficeOpenXml.Style.ExcelFillStyle]$TitleFillPattern='None',
        [Switch]$TitleBold,
        [Int]$TitleSize=22,
        [System.Drawing.Color]$TitleBackgroundColor,
        [String[]]$PivotRows,
        [String[]]$PivotColumns,
        $PivotData,
        [Switch]$PivotDataToColumn,
        [String]$Password,
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType='Pie',
        [Switch]$IncludePivotTable,
        [Switch]$IncludePivotChart,
        [Switch]$NoLegend,
        [Switch]$ShowCategory,
        [Switch]$ShowPercent,
        [Switch]$AutoSize,
        [Switch]$Show,
        [Switch]$NoClobber,
        [Switch]$FreezeTopRow,
        [Switch]$FreezeFirstColumn,
        [Switch]$FreezeTopRowFirstColumn,
        [Int[]]$FreezePane,
        [Switch]$AutoFilter,
        [Switch]$BoldTopRow,
        [Switch]$NoHeader,
        [String]$RangeName,
        [ValidateScript({
            if ($_.Contains(' ')) {
                throw 'Tablename has spaces.'
            }
            elseif (-not $_) {
                throw 'Tablename is null or empty.'
            }
            elseif ($_[0] -notmatch '[a-z]') {
                throw 'Tablename start with invalid character.'
            }
            else {
                $true
            }
        })] 
        [String]$TableName,
        [OfficeOpenXml.Table.TableStyles]$TableStyle='Medium6',
        [Object[]]$ConditionalFormat,
        [Object[]]$ConditionalText,
        [Object[]]$ExcelChartDefinition,
        [ScriptBlock]$CellStyleSB,
        [String[]]$HideSheet,
        [Switch]$KillExcel,
        [Switch]$AutoNameRange,
        $StartRow=1,
        $StartColumn=1,
        [Switch]$PassThru,
        [String]$Numberformat='General',
        [Switch]$Now
    )

    Begin {
    	$script:Header = $null
        if ($KillExcel) {
            Get-Process excel -ErrorAction Ignore | Stop-Process
            while (Get-Process excel -ErrorAction Ignore) {}
        }

        Try {
            if ($Now) {
                $Path=[System.IO.Path]::GetTempFileName() -replace '\.tmp','.xlsx'                
                $Show=$true
                $AutoSize=$true
                $AutoFilter=$true
            }

            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)

            if (Test-Path $path) {
                Write-Debug "File `"$Path`" already exists"
            }
            $pkg = New-Object OfficeOpenXml.ExcelPackage $Path

            $ws  = $pkg | Add-WorkSheet -WorkSheetname $WorkSheetname -NoClobber:$NoClobber

            foreach ($format in $ConditionalFormat ) {
                $target = "Add$($format.Formatter)"
                $rule = ($ws.ConditionalFormatting).PSObject.Methods[$target].Invoke($format.Range, $format.IconType)
                $rule.Reverse = $format.Reverse
            }

            # Force at least one cell value
            #$ws.Cells[1, 1].Value = ''


            $Row = $StartRow
            if ($Title) {
                $ws.Cells[$Row, $StartColumn].Value = $Title

                $ws.Cells[$Row, $StartColumn].Style.Font.Size = $TitleSize
                if ($TitleBold) {
                    #set title to Bold if -TitleBold was specified.
                    #Otherwise the default will be unbolded.
                    $ws.Cells[$Row, $StartColumn].Style.Font.Bold = $True
                }
                $ws.Cells[$Row, $StartColumn].Style.Fill.PatternType = $TitleFillPattern

                #can only set TitleBackgroundColor if TitleFillPattern is something other than None
                if ($TitleBackgroundColor -AND ($TitleFillPattern -ne 'None')) {
                    $ws.Cells[$Row, $StartColumn].Style.Fill.BackgroundColor.SetColor($TitleBackgroundColor)
                }
                else {
                    Write-Warning "Title Background Color ignored. You must set the TitleFillPattern parameter to a value other than 'None'. Try 'Solid'."
                }

                $Row += 1
            }

        } 
        Catch {
            if ($AlreadyExists) {
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
        if ($firstTimeThru) {
            $firstTimeThru=$false
            $isDataTypeValueType = $TargetData.GetType().name -match "string|bool|byte|char|decimal|double|float|int|long|sbyte|short|uint|ulong|ushort"
            Write-Verbose "DataTypeName is '$($TargetData.GetType().name)' isDataTypeValueType $isDataTypeValueType"
        }

        if ($isDataTypeValueType) {
            $ColumnIndex = $StartColumn

            $targetCell = $ws.Cells[$Row, $ColumnIndex]

            $r=$null
            $cellValue=$TargetData
            if ([Double]::TryParse([string]$cellValue,[System.Globalization.NumberStyles]::Any,[System.Globalization.NumberFormatInfo]::CurrentInfo, [ref]$r)) {
                $targetCell.Value = $r
                $targetCell.Style.Numberformat.Format=$Numberformat
            } else {
                $targetCell.Value = $cellValue
            }

            switch ($TargetData.$Name) {
                {$_ -is [datetime]} {$targetCell.Style.Numberformat.Format = "m/d/yy h:mm"}
            }

            $ColumnIndex += 1
            $Row += 1

        } else {
            if (!$script:Header) {

                $ColumnIndex = $StartColumn

                $script:Header = $TargetData.psobject.properties.name

                if ($NoHeader) {
                    # Don't push the headers to the spread sheet
                    $Row -= 1
                } else {
                    foreach ($Name in $script:Header) {
                        Write-Verbose "Add header '$Name'"
                        $ws.Cells[$Row, $ColumnIndex].Value = $Name
                        $ColumnIndex += 1
                    }
                }
            }

            $Row += 1
            $ColumnIndex = $StartColumn

            foreach ($Name in $script:Header) {
                $targetCell = $ws.Cells[$Row, $ColumnIndex]

                $cellValue=$TargetData.$Name

                if ($cellValue -is [string] -and $cellValue.StartsWith('=')) {
                    $targetCell.Formula = $cellValue
                } else {
                    $r=$null
                    if ([Double]::TryParse([string]$cellValue,[System.Globalization.NumberStyles]::Any,[System.Globalization.NumberFormatInfo]::CurrentInfo, [ref]$r)) {
                        $targetCell.Value = $r
                        $targetCell.Style.Numberformat.Format=$Numberformat
                        Write-Verbose "Add cell value '$r' in Numberformat '$Numberformat'"
                    } else {
                        $targetCell.Value = $cellValue
                        Write-Verbose "Add cell value '$cellValue' as String"
                    }
                }

                switch ($TargetData.$Name) {
                    {$_ -is [datetime]} {
                        $targetCell.Style.Numberformat.Format = "m/d/yy h:mm"
                    }
                }

                #[ref]$uriResult=$null
                #if ([uri]::TryCreate($cellValue, [System.UriKind]::Absolute, $uriResult)) {

                #    $targetCell.Hyperlink = [uri]$cellValue

                #    $namedStyle=$ws.Workbook.Styles.NamedStyles | where {$_.Name -eq 'HyperLink'}
                #    if (!$namedStyle) {
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
        Try {
            if ($AutoNameRange) {
                $totalRows=$ws.Dimension.Rows
                $totalColumns=$ws.Dimension.Columns

                foreach($c in 0..($totalColumns-1)) {
                    $targetRangeName = "$($script:Header[$c])"                

                    $targetColumn = $c+1
                    $theCell = $ws.Cells[2,$targetColumn,$totalRows,$targetColumn ]
                    $ws.Names.Add($targetRangeName, $theCell) | Out-Null

                    if ([OfficeOpenXml.FormulaParsing.ExcelUtilities.ExcelAddressUtil]::IsValidAddress($targetRangeName)) {
                        Write-Warning "AutoNameRange: Property name '$targetRangeName' is also a valid Excel address and may cause issues. Consider renaming the property name."
                    }
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
                $cec=$script:Header.Count

                $targetRange=$ws.Cells[$csr, $csc, $cer,$cec]

                $tbl = $ws.Tables.Add($targetRange, $TableName)

                $tbl.TableStyle=$TableStyle
            }

            if ($IncludePivotTable) {
                $pivotTableName = $WorkSheetname + "PivotTable"
                $wsPivot = $pkg | Add-WorkSheet -WorkSheetname $pivotTableName -NoClobber:$NoClobber

                $wsPivot.View.TabSelected = $true

                $pivotTableDataName=$WorkSheetname + "PivotTableData"

                if ($Title) {$startAddress="A2"}
                $pivotTable = $wsPivot.PivotTables.Add($wsPivot.Cells["A1"], $ws.Cells[$dataRange], $pivotTableDataName)

                if ($PivotRows) {
                    foreach ($Row in $PivotRows) {
                        $null=$pivotTable.RowFields.Add($pivotTable.Fields[$Row])
                    }
                }

                if ($PivotColumns) {
                    foreach ($Column in $PivotColumns) {
                        $null=$pivotTable.ColumnFields.Add($pivotTable.Fields[$Column])
                    }
                }

                if ($PivotData) {
                    if ($PivotData -is [hashtable] -or $PivotData -is [System.Collections.Specialized.OrderedDictionary]) {
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
                    if ($PivotDataToColumn) {
                        $pivotTable.DataOnRows = $false
                    }
                }

                if ($IncludePivotChart) {
                    $chart = $wsPivot.Drawings.AddChart("PivotChart", $ChartType, $pivotTable)

                    $chart.DataLabel.ShowCategory=$ShowCategory
                    $chart.DataLabel.ShowPercent=$ShowPercent

                    if ($NoLegend) { $chart.Legend.Remove() }

                    $chart.SetPosition(1, 0, 6, 0)
                    $chart.SetSize(600, 400)
                }
            }

            if ($Password) { $ws.Protection.SetPassword($Password) }

            if ($AutoFilter) {
                $ws.Cells[$dataRange].AutoFilter=$true
            }

            if ($FreezeTopRow) { $ws.View.FreezePanes(2,1) }
            if ($FreezeTopRowFirstColumn) { $ws.View.FreezePanes(2,2) }
            if ($FreezeFirstColumn) { $ws.View.FreezePanes(1,2) }

            if ($FreezePane) {
                $freezeRow,$freezeColumn=$FreezePane
                if (!$freezeColumn -or $freezeColumn -eq 0) {$freezeColumn=1}

                if ($freezeRow -gt 1) {
                    $ws.View.FreezePanes($freezeRow,$freezeColumn)
                }
            }

            if ($BoldTopRow) {
                $range=$ws.Dimension.Address -replace $ws.Dimension.Rows, "1"
                $ws.Cells[$range].Style.Font.Bold=$true
            }

            if ($AutoSize) { $ws.Cells.AutoFitColumns() }

            #$pkg.Workbook.View.ActiveTab = $ws.SheetID

            foreach ($Sheet in $HideSheet) {
                $pkg.Workbook.WorkSheets[$Sheet].Hidden="Hidden"
            }

            $chartCount=0
            foreach ($chartDef in $ExcelChartDefinition) {
                $ChartName = "Chart"+(Split-Path -Leaf ([System.IO.path]::GetTempFileName())) -replace 'tmp|\.',''
                $chart = $ws.Drawings.AddChart($ChartName, $chartDef.ChartType)
                $chart.Title.Text = $chartDef.Title

                if ($chartDef.NoLegend) {
                    $chart.Legend.Remove()
                }
                #$chart.Datalabel.ShowLegendKey = $true
            
                if ($chart.Datalabel -ne $null) {
                    $chart.Datalabel.ShowCategory  = $chartDef.ShowCategory
                    $chart.Datalabel.ShowPercent   = $chartDef.ShowPercent
                }

                $chart.SetPosition($chartDef.Row, $chartDef.RowOffsetPixels,$chartDef.Column, $chartDef.ColumnOffsetPixels)
                $chart.SetSize($chartDef.Width, $chartDef.Height)

                $chartDefCount = @($chartDef.YRange).Count
                if ($chartDefCount -eq 1) {
                    $Series=$chart.Series.Add($chartDef.YRange, $chartDef.XRange)
                
                    $SeriesHeader=$chartDef.SeriesHeader
                    if (!$SeriesHeader) {$SeriesHeader="Series 1"}

                    $Series.Header = $SeriesHeader
                } else {
                    for($idx=0; $idx -lt $chartDefCount; $idx+=1) {
                        $Series=$chart.Series.Add($chartDef.YRange[$idx], $chartDef.XRange)                    

                        if ($chartDef.SeriesHeader.Count -gt 0) {
                            $SeriesHeader=$chartDef.SeriesHeader[$idx]
                        }
                    
                        if (!$SeriesHeader) {$SeriesHeader="Series $($idx)"}

                        $Series.Header = $SeriesHeader
                        $SeriesHeader=$null
                    }
                }
            }

            if ($ConditionalText) {
                foreach ($targetConditionalText in $ConditionalText) {
                    $target = "Add$($targetConditionalText.ConditionalType)"

                    $Range=$targetConditionalText.Range
                    if (!$Range) { $Range=$ws.Dimension.Address }

                    $rule=($ws.Cells[$Range].ConditionalFormatting).PSObject.Methods[$target].Invoke()

                    if ($targetConditionalText.Text) {                
                        if ($targetConditionalText.ConditionalType -match "equal|notequal|lessthan|lessthanorequal|greaterthan|greaterthanorequal") {
                            $rule.Formula= $targetConditionalText.Text
                        } else {
                            $rule.Text = $targetConditionalText.Text
                        }
                    }

                    $rule.Style.Font.Color.Color = $targetConditionalText.ConditionalTextColor
                    $rule.Style.Fill.PatternType=$targetConditionalText.PatternType
                    $rule.Style.Fill.BackgroundColor.Color=$targetConditionalText.BackgroundColor
               }
            }

            if ($CellStyleSB) {
                $TotalRows=$ws.Dimension.Rows
                $LastColumn=(Get-ExcelColumnName $ws.Dimension.Columns).ColumnName
                & $CellStyleSB $ws $TotalRows $LastColumn
            }

            if ($PassThru) {
                $pkg
            } else {
                $pkg.Save()
                $pkg.Dispose()

                if ($Show) {Invoke-Item $Path}
            }
        }
        Catch {
            throw "Failed exporting the worksheet: $_"
        }
    }
}
