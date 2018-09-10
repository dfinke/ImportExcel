$path = "$Env:TEMP\test.xlsx"
remove-item -path $path -ErrorAction SilentlyContinue

#Export some sales data to Excel, format it as a table and put a data-bar in

$excel =  ConvertFrom-Csv    @"
Product, City, Gross, Net
Apple, London , 300, 250
Orange, London , 400, 350
Banana, London , 300, 200
Orange, Paris,   600, 500
Banana, Paris,   300, 200
Apple, New York, 1200,700
"@  | Export-Excel  -Path $path  -TableStyle Medium13 -tablename "RawData" -ConditionalFormat @{Range="C2:C7"; DataBarColor="Green"}  -PassThru

#Add a pivot table on the same sheet, using this data. set the table style and number format. Use the "City" as row names, and "Product" for columnnames, and total both the gross and net columns
#Add a pivot chart (defined in a hash table)
Add-PivotTable -Address $excel.Sheet1.Cells["F1"] -SourceWorkSheet $Excel.Sheet1 -SourceRange $Excel.Sheet1.Dimension.Address -PivotTableName "Sales" -PivotTableSyle "Medium12" -Activate  `
                 -PivotRows "City" -PivotColumns "Product" -PivotData @{Gross="Sum";Net="Sum"} -PivotNumberFormat "$#,##0.00" -PivotTotals "Both"  -PivotChartDefinition @{
                        Title="Gross and net by city and product";
                        ChartType="ColumnClustered";
                        Column=11; Width=500; Height=360;
                        YMajorUnit=500; YMinorUnit=100; YAxisNumberformat="$#,##0"
                        LegendPostion="Bottom"}
#Save and open in excel
Close-ExcelPackage $excel -Show