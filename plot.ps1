[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification="False positives")]
param()
class PSPlot {
    hidden $path
    hidden $pkg
    hidden $ws
    hidden $chart

    PSPlot() {
        $this.path=[System.IO.Path]::GetTempFileName() -replace "\.tmp", ".xlsx"
        $this.pkg = New-Object OfficeOpenXml.ExcelPackage $this.path
        $this.ws=$this.pkg.Workbook.Worksheets.Add("plot")
    }

    [PSPlot] Plot($yValues) {

        $this.NewChart()

        $xValues = 0..$yValues.Count

        $xCol = 'A'
        $yCol = 'B'

        $this.AddDataToSheet($xCol,$yCol,'x','y',$xValues,$yValues)
        $this.AddSeries($xCol,$yCol,$yValues)
        $this.SetChartPosition($yCol)

        return $this
    }

    [PSPlot] Plot($yValues,[string]$options) {
        $this.NewChart()

        $xValues = 0..$yValues.Count

        $xCol = 'A'
        $yCol = 'B'

        $this.AddDataToSheet($xCol,$yCol,'x','y',$xValues,$yValues)
        $this.AddSeries($xCol,$yCol,$yValues)

        $this.SetMarkerInfo($options)
        $this.SetChartPosition($yCol)

        return $this
    }

    [PSPlot] Plot($xValues,$yValues) {

        $this.NewChart()

        $xCol = 'A'
        $yCol = 'B'

        $this.AddDataToSheet($xCol,$yCol,'x','y',$xValues,$yValues)
        $this.AddSeries($xCol,$yCol,$yValues)

        $this.SetChartPosition($yCol)

        return $this
    }

    [PSPlot] Plot($xValues,$yValues,[string]$options) {
        $this.NewChart()

        $xCol = 'A'
        $yCol = 'B'

        $this.AddDataToSheet($xCol,$yCol,'x','y',$xValues,$yValues)
        $this.AddSeries($xCol,$yCol,$yValues)

        $this.SetMarkerInfo($options)

        $this.SetChartPosition($yCol)

        return $this
    }

    [PSPlot] Plot($xValues,$yValues,$x1Values,$y1Values) {

        $this.NewChart()

        $xCol = 'A'
        $yCol = 'B'

        $this.AddDataToSheet($xCol,$yCol,'x','y',$xValues,$yValues)
        $this.AddSeries($xCol,$yCol,$yValues)

        $xCol=$this.GetNextColumnName($yCol)
        $yCol=$this.GetNextColumnName($xCol)

        $this.AddDataToSheet($xCol,$yCol,'x1','y1',$x1Values,$y1Values)
        $this.AddSeries($xCol,$yCol,$y1Values)

        $this.SetChartPosition($yCol)

        return $this
    }

    [PSPlot] Plot($xValues,$yValues,$x1Values,$y1Values,$x2Values,$y2Values) {

        $this.NewChart()

        $xCol = 'A'
        $yCol = 'B'

        $this.AddDataToSheet($xCol,$yCol,'x','y',$xValues,$yValues)
        $this.AddSeries($xCol,$yCol,$yValues)

        $xCol=$this.GetNextColumnName($yCol)
        $yCol=$this.GetNextColumnName($xCol)

        $this.AddDataToSheet($xCol,$yCol,'x1','y1',$x1Values,$y1Values)
        $this.AddSeries($xCol,$yCol,$y1Values)

        $xCol=$this.GetNextColumnName($yCol)
        $yCol=$this.GetNextColumnName($xCol)

        $this.AddDataToSheet($xCol,$yCol,'x2','y2',$x2Values,$y2Values)
        $this.AddSeries($xCol,$yCol,$y2Values)

        $this.SetChartPosition($yCol)

        return $this
    }

    [PSPLot] SetChartPosition($yCol) {
        $columnNumber = $this.GetColumnNumber($yCol)+1
        $this.chart.SetPosition(1,0,$columnNumber,0)

        return $this
    }

    AddSeries($xCol,$yCol,$yValues) {
        $yRange = "{0}2:{0}{1}" -f $yCol,($yValues.Count+1)
        $xRange = "{0}2:{0}{1}" -f $xCol,($yValues.Count+1)
        $Series=$this.chart.Series.Add($yRange,$xRange)
    }

    hidden SetMarkerInfo([string]$options) {
        $c=$options.Substring(0,1)
        $m=$options.Substring(1)

        $cmap=@{r='red';g='green';b='blue';i='indigo';v='violet';c='cyan'}
        $mmap=@{Ci='Circle';Da='Dash';di='diamond';do='dot';pl='plus';sq='square';tr='triangle'}

        $this.chart.Series[0].Marker = $mmap.$m
        $this.chart.Series[0].MarkerColor = $cmap.$c
        $this.chart.Series[0].MarkerLineColor = $cmap.$c
    }

    hidden [string]GetNextColumnName($columnName) {
        return $this.GetColumnName($this.GetColumnNumber($columnName)+1)
    }

    hidden [int]GetColumnNumber($columnName) {
        $sum=0

        $columnName.ToCharArray() |
            ForEach-Object {
                $sum*=26
                $sum+=[char]$_.tostring().toupper()-[char]'A'+1
            }

        return $sum
    }

    hidden [string]GetColumnName($columnNumber) {
        $dividend = $columnNumber
        $columnName = @()
        while($dividend -gt 0) {
            $modulo      = ($dividend - 1) % 26
            $columnName += [char](65 + $modulo)
            $dividend    = [int](($dividend -$modulo)/26)
        }

        return ($columnName -join '')
    }

    hidden AddDataToSheet($xColumn,$yColumn,$xHeader,$yHeader,$xValues,$yValues) {
        $count=$yValues.Count
        $this.ws.Cells["$($xColumn)1"].Value=$xHeader
        $this.ws.Cells["$($yColumn)1"].Value=$yHeader

        for ($idx= 0; $idx-lt $count; $idx++) {
            $row=$idx+2
            $this.ws.Cells["$($xColumn)$($row)"].Value=$xValues[$idx]
            $this.ws.Cells["$($yColumn)$($row)"].Value=$yValues[$idx]
        }
    }

    hidden NewChart() {
        $chartType="XYScatter"
        #$chartType="line"
        $this.chart=$this.ws.Drawings.AddChart("plot", $chartType)
        $this.chart.Title.Text = 'Plot'
        $this.chart.Legend.Remove()
        $this.SetChartSize(300,300)
    }

    [PSPlot] SetChartSize([int]$width,[int]$height){
        $this.chart.SetSize($width, $height)

        return $this
    }

    [PSPlot] Title($title) {
        $this.chart.Title.Text = $title

        return $this
    }

    Show() {
        $this.pkg.Save()
        $this.pkg.Dispose()
        Invoke-Item $this.path
    }
}