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
        [string[]]$PivotRows,
        [string[]]$PivotColumns,
        [string[]]$PivotData,
        [ValidateSet("Area3D","AreaStacked3D","AreaStacked1003D","BarClustered3D","BarStacked3D","BarStacked1003D","Column3D","ColumnClustered3D","ColumnStacked3D","ColumnStacked1003D","Line3D","Pie3D","PieExploded3D","Area","AreaStacked","AreaStacked100","BarClustered","BarOfPie","BarStacked","BarStacked100","Bubble","Bubble3DEffect","ColumnClustered","ColumnStacked","ColumnStacked100","ConeBarClustered","ConeBarStacked","ConeBarStacked100","ConeCol","ConeColClustered","ConeColStacked","ConeColStacked100","CylinderBarClustered","CylinderBarStacked","CylinderBarStacked100","CylinderCol","CylinderColClustered","CylinderColStacked","CylinderColStacked100","Doughnut","DoughnutExploded","Line","LineMarkers","LineMarkersStacked","LineMarkersStacked100","LineStacked","LineStacked100","Pie","PieExploded","PieOfPie","PyramidBarClustered","PyramidBarStacked","PyramidBarStacked100","PyramidCol","PyramidColClustered","PyramidColStacked","PyramidColStacked100","Radar","RadarFilled","RadarMarkers","StockHLC","StockOHLC","StockVHLC","StockVOHLC","Surface","SurfaceTopView","SurfaceTopViewWireframe","SurfaceWireframe","XYScatter","XYScatterLines","XYScatterLinesNoMarkers","XYScatterSmooth","XYScatterSmoothNoMarkers")]
        $ChartType="Pie",
        [Switch]$IncludePivotTable,
        [Switch]$IncludePivotChart,
        [Switch]$AutoFitColumns,
        [Switch]$Show,
        [Switch]$Force
    )

    Begin {

        if(Test-Path $Path) {
            if($Force) {
                Remove-Item $Path
            } else {
                throw "$Path already exists"
            }
        }

        $pkg = New-Object OfficeOpenXml.ExcelPackage $Path
        $ws = $pkg.Workbook.Worksheets.Add("Sheet1")
        $Row = 1
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

        if($AutoFitColumns) { $ws.Cells.AutoFitColumns()}

        if($IncludePivotTable) {

            $wsPivot = $pkg.Workbook.Worksheets.Add("PivotTable1")
            $wsPivot.View.TabSelected = $true

            $range="{0}:{1}" -f $ws.Dimension.Start.Address, $ws.Dimension.End.Address
            $pivotTable = $wsPivot.PivotTables.Add($wsPivot.Cells["A1"], $ws.Cells[$range], "PivotTableData")

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
                #$ChartType="Pie"
                #$ChartType="PieExploded3D"
                $chart = $wsPivot.Drawings.AddChart("PivotChart", $ChartType, $pivotTable)
                $chart.SetPosition(1, 0, 6, 0)
                $chart.SetSize(600, 400)
            }
        }

        $pkg.Save()
        $pkg.Dispose()

        if($Show) {Invoke-Item $Path}
    }
}