function DoChart {
     param(
         $targetData,
         $title,
         $chartType
     )

     if($targetData[0] -is [System.ValueType]) {
         $chart = New-ExcelChart -YRange "A1:A$($targetData.count)" -Title $title -ChartType $chartType
     } else {
         $xyRange = Get-XYRange $targetData

         $X = $xyRange.XRange.ExcelColumn
         $XRange = "{0}2:{0}{1}" -f $X,($targetData.count+1)

         $Y = $xyRange.YRange.ExcelColumn
         $YRange = "{0}2:{0}{1}" -f $Y,($targetData.count+1)

         $chart = New-ExcelChart -XRange $xRange -YRange $yRange -Title $title -ChartType $chartType
     }

     $xlFile = (New-TemporaryFile).fullname -replace "tmp","xlsx"
     $targetData | Export-Excel $xlFile -ExcelChartDefinition $chart -Show -AutoSize

 }

 function BarChart {
     param($targetData,$title)

     DoChart $targetData $title BarStacked
 }

 function PieChart {
     param($targetData,$title)

     DoChart $targetData $title Pie
 }

 function LineChart {
     param($targetData,$title)

     DoChart $targetData $title Line
 }

 function ColumnChart {
     param($targetData,$title)

     DoChart $targetData $title ColumnStacked
 }
