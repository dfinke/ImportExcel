try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$sql           = @"
  SELECT     rootfile.baseName     ,  rootfile.extension   , Image.fileWidth         AS width             ,  image.fileHeight AS height ,
             metadata.dateDay      ,  metadata.dateMonth   , metadata.dateYear       ,  Image.captureTime AS dateTaken,
             metadata.hasGPS       ,  metadata.GPSLatitude , metadata.GPSLongitude   ,
             metadata.focalLength  ,  metadata.flashFired  , metadata.ISOSpeedRating AS ISOSpeed,
             metadata.Aperture     AS apertureValue        , metadata.ShutterSpeed AS shutterSpeedValue,
             Image.bitdepth        ,  image.colorLabels    ,
             Camera.Value          AS cameraModel          , LensRef.value          AS lensModel
      FROM   Adobe_images                  image
      JOIN   AgLibraryFile              rootFile  ON   rootfile.id_local =      image.rootFile
      JOIN   AgharvestedExifMetadata    metadata  ON      image.id_local =   metadata.image
 LEFT JOIN   AgInternedExifLens          LensRef  ON    LensRef.id_Local =   metadata.lensRef
 LEFT JOIN   AgInternedExifCameraModel    Camera  ON     Camera.id_local =  metadata.cameraModelRef
"@

#Sql Statement gets 20 columns of data from Adobe lightroom database
#Define a pivot table and chart for total pictures with each lens.

$pt = @{"LensPivot" = @{ "PivotTableName"    =   "LensPivot";
                         "SourceWorkSheet"   =   "Sheet1" ;
                         "PivotRows"         =   "LensModel" ;
                         "PivotData"         = @{"basename" = "Count"}  ;
                         "IncludePivotChart" =    $true ;
                         "NoLegend"          =    $true ;
                         "ShowPercent"       =    $true ;
                         "ChartType"         =   "Pie"  ;
                         "ChartTitle"        =   "Split by Lens" }
}

#we want to add 3 columns, translate Apperture value and Shutter speed value into familar f/ and seconds notation, and use these and ISO to calculate EV level
$Avalue        = {"=IF(P$ROW>6.63,TEXT(ROUND(Sqrt(Power(2,O$ROW)),1),`"`"`"f/`"`"0.0`")," +
                          "TEXT(ROUND(Sqrt(Power(2,O$ROW)),1),`"`"`"f/`"`"0.0`"))"}
$Svalue        = {"=IF(P$ROW>2,TEXT(ROUND(POWER(2,P$ROW),0),`"`"`"1/`"`"0`"`"sec`"`"`"),"+
                        "IF(P$ROW>3.32,TEXT(ROUND(1/POWER(2,P$ROW),2),`"0.0`"`"Sec`"`"`"),"+
                            "TEXT(ROUND(1/POWER(2,P$ROW),2),`"0`"`"Sec`"`"`")))"}
$evValue       = {"=ROUND(P$Row+O$Row-(LOG(N$Row/100,2)),0)" }

#remove and recreate the file
Remove-Item -Path "~\Documents\temp.xlsx" -ErrorAction SilentlyContinue

#Open a connection to the ODBC source "LR" (which points to the SQLLite DB for Lightroom), run the SQL query, and drop into Excel - in sheet1, autosizing columns.
$e = Send-SQLDataToExcel   -Path "~\Documents\temp.xlsx" -WorkSheetname "Sheet1" -Connection "DSN=LR" -SQL $sql -AutoSize -Passthru

#Add columns, then format them and hide the ones which aren't of interest.
Set-ExcelColumn   -Worksheet    $e.workbook.Worksheets["sheet1"] -Column 21  -Value $Avalue  -Heading "Apperture"
Set-ExcelColumn   -Worksheet    $e.workbook.Worksheets["sheet1"] -Column 22  -Value $Svalue  -Heading "Shutter"
Set-ExcelColumn   -Worksheet    $e.workbook.Worksheets["sheet1"] -Column 23  -Value $Evvalue -Heading "Ev"
Set-ExcelRange   -Address      $e.workbook.Worksheets["sheet1" ].Column(21) -HorizontalAlignment Left  -AutoFit
Set-ExcelRange   -Address      $e.workbook.Worksheets["sheet1" ].Column(22) -HorizontalAlignment Right -AutoFit
@(5,6,7,13,15,16,17,18) | ForEach-Object {
  Set-ExcelRange -Address      $e.workbook.Worksheets["sheet1" ].Column($_) -Hidden
}

#Center the column labels.
Set-ExcelRange   -Address      $e.workbook.Worksheets["sheet1" ].Row(1)     -HorizontalAlignment Center

#Format the data as a nice Table, Create the pivot table & chart defined above, show the file in Excel in excel after saving.
Export-Excel -ExcelPackage $e      -WorkSheetname "sheet1"  -TableName  "Table" -PivotTableDefinition $pt -Show

############################################################

Remove-Item  .\demo3.xlsx
#Database query to get race wins, Poles and fastest lapes for the 25 best drivers; we already have a connection to the DB in $dbSessions
$session =  $DbSessions["f1"]
$SQL     =  @"
    SELECT TOP 25 DriverName,
                  Count(RaceDate)   AS Races,
                  Count(Win)        AS Wins,
                  Count(Pole)       AS Poles,
                  Count(FastestLap) AS Fastlaps
    FROM  Results
    GROUP BY DriverName
    ORDER BY (Count(win)) DESC
"@

#Run the query and put the results in workshet "Winners", autosize the columns and hold on to the ExcelPackage object
$Excel   =  Send-SQLDataToExcel -SQL $sql -Session $session -path .\demo3.xlsx -WorkSheetname "Winners"   -AutoSize -Passthru
#Create and format columns for the ratio of Wins to poles and fast laps.
Set-ExcelColumn   -ExcelPackage $Excel -WorkSheetname "Winners" -column 6  -Heading "WinsToPoles" -Value {"=D$row/C$row"}
Set-ExcelColumn   -ExcelPackage $Excel -WorkSheetname "Winners" -column 7  -Heading "WinsToFast"  -Value {"=E$row/C$row"}
6..7 | ForEach-Object {
    Set-ExcelRange -Address    $Excel.Workbook.Worksheets["Winners"].column($_) -NumberFormat "0.0%" -AutoFit }
#Define a chart to show the relationship of lest on an XY Grid, create the ranges required in the, add the chart and show the file in Excel in excel after saving.
$chart = New-ExcelChartDefinition -NoLegend -ChartType XYScatter -XRange WinsToFast -YRange WinsToPoles -ShowCategory -Column 7 -Width 2000 -Height 700
Export-Excel -ExcelPackage $Excel -WorkSheetname "Winners" -AutoNameRange -ExcelChartDefinition $chart -Show

