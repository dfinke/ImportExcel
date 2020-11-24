

Describe "Creating workbook with a single line" {
    BeforeAll {
        $path = "TestDrive:\test.xlsx"
        remove-item -path $path -ErrorAction SilentlyContinue
        ConvertFrom-Csv    @"
Product, City, Gross, Net
Apple, London , 300, 250
Orange, London , 400, 350
Banana, London , 300, 200
Orange, Paris,   600, 500
Banana, Paris,   300, 200
Apple, New York, 1200,700

"@  | Export-Excel  -Path $path  -TableStyle Medium13 -tablename "RawData" -ConditionalFormat @{Range = "C2:C7"; DataBarColor = "Green" } -ExcelChartDefinition @{ChartType = "Doughnut"; XRange = "A2:B7"; YRange = "C2:C7"; width = 800; }  -PivotTableDefinition @{Sales = @{
                PivotRows = "City"; PivotColumns = "Product"; PivotData = @{Gross = "Sum"; Net = "Sum" }; PivotNumberFormat = "$#,##0.00"; PivotTotals = "Both"; PivotTableStyle = "Medium12"; Activate = $true

                PivotChartDefinition = @{Title = "Gross and net by city and product"; ChartType = "ColumnClustered"; Column = 6; Width = 600; Height = 360; YMajorUnit = 500; YMinorUnit = 100; YAxisNumberformat = "$#,##0"; LegendPosition = "Bottom" }
            }
        }

        $excel = Open-ExcelPackage $path
        $ws1 = $excel.Workbook.Worksheets[1]
        $ws2 = $excel.Workbook.Worksheets[2]
    }
    Context "Data Page" {
        It "Inserted the data and created the table                                                " {
            $ws1.Tables[0]                                              | Should -Not -BeNullOrEmpty
            $ws1.Tables[0].Address.Address                              | Should      -Be "A1:D7"
            $ws1.Tables[0].StyleName                                    | Should      -Be "TableStyleMedium13"
        }
        It "Applied conditional formatting                                                         " {
            $ws1.ConditionalFormatting[0]                               | Should -Not -BeNullOrEmpty
            $ws1.ConditionalFormatting[0].type.ToString()               | Should      -Be "DataBar"
            $ws1.ConditionalFormatting[0].Color.G                       | Should      -BeGreaterThan 100
            $ws1.ConditionalFormatting[0].Color.R                       | Should      -BeLessThan    100
            $ws1.ConditionalFormatting[0].Address.Address               | Should      -Be "C2:C7"
        }
        It "Added the chart                                                                        " {
            $ws1.Drawings[0]                                            | Should -Not -BeNullOrEmpty
            $ws1.Drawings[0].ChartType.ToString()                       | Should      -Be "DoughNut"
            $ws1.Drawings[0].Series[0].Series                           | Should      -Be "'Sheet1'!C2:C7"
        }
    }
    Context "PivotTable" {
        it "Created the PivotTable on a new page                                                   " {
            $ws2                                                        | Should -Not -BeNullOrEmpty
            $ws2.PivotTables[0]                                         | Should -Not -BeNullOrEmpty
            $ws2.PivotTables[0].Fields.Count                            | Should      -Be 4
            $ws2.PivotTables[0].DataFields[0].Format                    | Should      -Be "$#,##0.00"
            $ws2.PivotTables[0].RowFields[0].Name                       | Should      -Be "City"
            $ws2.PivotTables[0].ColumnFields[0].Name                    | Should      -Be "Product"
            $ws2.PivotTables[0].RowGrandTotals                          | Should      -Be $true
            $ws2.PivotTables[0].ColumGrandTotals                        | Should      -Be $true   #Epplus's mis-spelling of column not mine
        }
        it "Made the PivotTable page active                                                        " {
            Set-ItResult -Pending -Because "Bug in EPPLus 4.5"
            $ws2.View.TabSelected                                       | Should      -Be $true
        }
        it "Created the Pivot Chart                                                                " {
            $ws2.Drawings[0]                                            | Should -Not -BeNullOrEmpty
            $ws2.Drawings[0].ChartType.ToString()                       | Should      -Be ColumnClustered
            $ws2.Drawings[0].YAxis.MajorUnit                            | Should      -Be 500
            $ws2.Drawings[0].YAxis.MinorUnit                            | Should      -Be 100
            $ws2.Drawings[0].YAxis.Format                               | Should      -Be "$#,##0"
            $ws2.Drawings[0].Legend.Position.ToString()                 | Should      -Be "Bottom"
        }

    }

}
