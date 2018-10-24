

$path = "$Env:TEMP\test.xlsx"
remove-item -path $path -ErrorAction SilentlyContinue
ConvertFrom-Csv    @"
Product, City, Gross, Net
Apple, London , 300, 250
Orange, London , 400, 350
Banana, London , 300, 200
Orange, Paris,   600, 500
Banana, Paris,   300, 200
Apple, New York, 1200,700
 
"@  | Export-Excel  -Path $path  -TableStyle Medium13 -tablename "RawData" -ConditionalFormat @{Range="C2:C7"; DataBarColor="Green"} -ExcelChartDefinition @{ChartType="Doughnut";XRange="A2:B7"; YRange="C2:C7"; width=800; }  -PivotTableDefinition @{Sales=@{
            PivotRows="City"; PivotColumns="Product"; PivotData=@{Gross="Sum";Net="Sum"}; PivotNumberFormat="$#,##0.00"; PivotTotals="Both"; PivotTableStyle="Medium12"; Activate=$true
 
            PivotChartDefinition=@{Title="Gross and net by city and product"; ChartType="ColumnClustered"; Column=6; Width=600; Height=360; YMajorUnit=500; YMinorUnit=100; YAxisNumberformat="$#,##0"; LegendPosition="Bottom"}}}

$excel = Open-ExcelPackage $path
$ws1 = $excel.Workbook.Worksheets[1]
$ws2  = $excel.Workbook.Worksheets[2]
Describe "Creating workbook with a single line" {
    Context "Data Page" {
        It "Inserted the data and created the table                                                " {
            $ws1.Tables[0]                                              | Should not beNullOrEmpty
            $ws1.Tables[0].Address.Address                              | Should     be "A1:D7"
            $ws1.Tables[0].StyleName                                    | Should     be "TableStyleMedium13"
        }
        It "Applied conditional formatting                                                         " {
            $ws1.ConditionalFormatting[0]                               | Should not beNullOrEmpty
            $ws1.ConditionalFormatting[0].type.ToString()               | Should     be "DataBar"
            $ws1.ConditionalFormatting[0].Color.G                       | Should     beGreaterThan 100
            $ws1.ConditionalFormatting[0].Color.R                       | Should     beLessThan    100
            $ws1.ConditionalFormatting[0].Address.Address               | Should     be "C2:C7"
        }
        It "Added the chart                                                                        " {
            $ws1.Drawings[0]                                            | Should not beNullOrEmpty
            $ws1.Drawings[0].ChartType.ToString()                       | Should     be "DoughNut"
            $ws1.Drawings[0].Series[0].Series                           | Should     be "'Sheet1'!C2:C7"
        }
    }
    Context "PivotTable"    {
        it "Created the PivotTable on a new page and made it active                                " {
            $ws2                                                        | Should not beNullOrEmpty
            $ws2.PivotTables[0]                                         | Should not beNullOrEmpty
            $ws2.PivotTables[0].Fields.Count                            | Should     be 4
            $ws2.PivotTables[0].DataFields[0].Format                    | Should     be "$#,##0.00"
            $ws2.PivotTables[0].RowFields[0].Name                       | Should     be "City"
            $ws2.PivotTables[0].ColumnFields[0].Name                    | Should     be "Product"
            $ws2.PivotTables[0].RowGrandTotals                          | Should     be $true
            $ws2.PivotTables[0].ColumGrandTotals                        | Should     be $true   #Epplus's mis-spelling of column not mine
            $ws2.View.TabSelected                                       | Should     be $true
        }
        it "Created the Pivot Chart                                                                " {
            $ws2.Drawings[0]                                            | Should not beNullOrEmpty
            $ws2.Drawings[0].ChartType.ToString()                       | Should     be ColumnClustered
            $ws2.Drawings[0].YAxis.MajorUnit                            | Should     be 500
            $ws2.Drawings[0].YAxis.MinorUnit                            | Should     be 100
            $ws2.Drawings[0].YAxis.Format                               | Should     be "$#,##0"
            $ws2.Drawings[0].Legend.Position.ToString()                 | Should     be "Bottom"
        }

    }

}
