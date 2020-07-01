try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

#Export some sales data to Excel, format it as a table and put a data-bar in.  For this example we won't create the pivot table during the export
$excel =  ConvertFrom-Csv    @"
Product, City, Gross, Net
Apple, London , 300, 250
Orange, London , 400, 350
Banana, London , 300, 200
Orange, Paris,   600, 500
Banana, Paris,   300, 200
Apple, New York, 1200,700
"@  | Export-Excel -PassThru -Path $xlSourcefile -TableStyle Medium13 -tablename "RawData" -ConditionalFormat @{Range="C2:C7"; DataBarColor="Green"}

#Add a pivot table, specify its address to put it on the same sheet, use the data that was just exported  set the table style and number format.
#Use the "City" for the row names, and "Product" for the columnnames, and sum both the gross and net values for each City/Product combination; add grand totals to rows and columns.
# activate the sheet and add a pivot chart (defined in a hash table)
Add-PivotTable -Address $excel.Sheet1.Cells["F1"] -SourceWorkSheet $Excel.Sheet1 -SourceRange $Excel.Sheet1.Dimension.Address -PivotTableName "Sales" -PivotTableStyle "Medium12"  -PivotNumberFormat "$#,##0.00"   `
                 -PivotRows "City" -PivotColumns "Product" -PivotData @{Gross="Sum";Net="Sum"}-PivotTotals "Both"  -Activate  -PivotChartDefinition @{
                        Title="Gross and net by city and product";
                        ChartType="ColumnClustered";
                        Column=11; Width=500; Height=360;
                        YMajorUnit=500; YMinorUnit=100; YAxisNumberformat="$#,##0"
                        LegendPosition="Bottom"}
#Save and open in excel
Close-ExcelPackage $excel -Show