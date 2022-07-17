
Describe "Creating small named ranges with hyperlinks" {
    BeforeAll {
        $scriptPath = $PSScriptRoot
        $dataPath = Join-Path  -Path $scriptPath -ChildPath "First10Races.csv"
        $WarningAction = "SilentlyContinue"
        $path = "TestDrive:\Results.xlsx"
        Remove-Item -Path $path -ErrorAction SilentlyContinue
        #Read race results, and group by race name : export 1 row to get headers, leaving enough rows aboce to put in a link for each race
        $results = Import-Csv -Path $dataPath |
        Select-Object  Race, @{n = "Date"; e = { [datetime]::ParseExact($_.date, "dd/MM/yyyy", (Get-Culture)) } }, FinishPosition, Driver, GridPosition, Team, Points |
        Group-Object -Property RACE
        $topRow = $lastDataRow = 1 + $results.Count
        $excel = $results[0].Group[0] | Export-Excel -Path $path -StartRow $TopRow  -BoldTopRow -PassThru

        #export each group (race) below the last one, without headers, and create a range for each using the group name (Race)
        foreach ($r in $results) {
            $excel = $R.Group | Export-Excel -ExcelPackage $excel -NoHeader -StartRow ($lastDataRow + 1) -RangeName $R.Name -PassThru -AutoSize
            $lastDataRow += $R.Group.Count
        }
        $worksheet = $excel.Workbook.Worksheets[1]
        $columns = $worksheet.Dimension.Columns

        1..$columns | ForEach-Object { Add-ExcelName -Range $worksheet.cells[$topRow, $_, $lastDataRow, $_] }                                                                        #Test Add-Excel Name on its own (outside Export-Excel)

        $scwarnVar = $null
        Set-ExcelColumn -Worksheet $worksheet -StartRow $topRow -Heading "PlacesGained/Lost" `
            -Value "=GridPosition-FinishPosition" -AutoNameRange -WarningVariable scWarnVar -WarningAction SilentlyContinue                 #Test as many set column options as possible.
        $columns ++

        #create a table which covers all the data. And define a pivot table which uses the same address range.
        $table = Add-ExcelTable -PassThru  -Range  $worksheet.cells[$topRow, 1, $lastDataRow, $columns]  -TableName "AllResults" -TableStyle Light4 `
            -ShowHeader -ShowFilter -ShowColumnStripes -ShowRowStripes:$false -ShowFirstColumn:$false -ShowLastColumn:$false -ShowTotal:$false   #Test Add-ExcelTable outside Export-Excel with as many options as possible.
        $pt = New-PivotTableDefinition -PivotTableName Analysis -SourceWorkSheet   $worksheet -SourceRange $table.address.address -PivotRows Driver -PivotData @{Points = "SUM" } -PivotTotals None

        $cf = Add-ConditionalFormatting -Address  $worksheet.cells[$topRow, $columns, $lastDataRow, $columns] -ThreeIconsSet Arrows  -Passthru                               #Test using cells[r1,c1,r2,c2]
        $cf.Icon2.Type = $cf.Icon3.Type = "Num"
        $cf.Icon2.Value = 0
        $cf.Icon3.Value = 1
        Add-ConditionalFormatting -Address $worksheet.cells["FinishPosition"] -RuleType Equal    -ConditionValue 1 -ForeGroundColor  ([System.Drawing.Color]::Purple) -Bold -Priority 1 -StopIfTrue   #Test Priority and stopIfTrue and using range name
        Add-ConditionalFormatting -Address $worksheet.Cells["GridPosition"]   -RuleType ThreeColorScale -Reverse                                                           #Test Reverse
        $ct = New-ConditionalText -Text "Ferrari"
        $ct2 = New-ConditionalText -Range $worksheet.Names["FinishPosition"].Address -ConditionalType LessThanOrEqual -Text 3 -ConditionalText ([System.Drawing.Color]::Red) -Background ([System.Drawing.Color]::White)      #Test New-ConditionalText in shortest and longest forms.
        #Create links for each group name (race) and Export them so they start at Cell A1; create a pivot table with definition just created, save the file and open in Excel
        $excel = $results | ForEach-Object { (New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList "Sheet1!$($_.Name)" , "$($_.name) GP") } |                                     #Test Exporting Hyperlinks with display property.
        Export-Excel -ExcelPackage $excel -AutoSize -PivotTableDefinition $pt -Calculate   -ConditionalFormat $ct, $ct2  -PassThru                                        #Test conditional text rules in conditional format (orignally icon sets only )

        $null = Add-Worksheet -ExcelPackage $excel -WorksheetName "Points1"
        Add-PivotTable -PivotTableName "Points1" -Address $excel.Points1.Cells["A1"] -ExcelPackage $excel -SourceWorkSheet sheet1 -SourceRange $excel.Sheet1.Tables[0].Address.Address -PivotRows Driver, Date -PivotData @{Points = "SUM" }  -GroupDateRow Date -GroupDatePart Years, Months

        $null = Add-Worksheet -ExcelPackage $excel -WorksheetName "Places1"
        $newpt = Add-PivotTable -PivotTableName "Places1" -Address $excel.Places1.Cells["A1"] -ExcelPackage $excel -SourceWorkSheet sheet1 -SourceRange $excel.Sheet1.Tables[0].Address.Address -PivotRows Driver, FinishPosition -PivotData @{Date = "Count" }  -GroupNumericRow FinishPosition -GroupNumericMin 1 -GroupNumericMax 25 -GroupNumericInterval 3 -PassThru
        $newpt.RowFields[0].SubTotalFunctions = [OfficeOpenXml.Table.PivotTable.eSubTotalFunctions]::None
        Close-ExcelPackage -ExcelPackage $excel

        $excel = Open-ExcelPackage $path
        $sheet = $excel.Workbook.Worksheets[1]
        $m = $results | Measure-Object -sum -Property count
        $expectedRows = 1 + $m.count + $m.sum
    }
    Context "Creating hyperlinks" {
        it "Put the data into the sheet and created the expected named ranges                      " {
            $sheet.Dimension.Rows                                       | Should      -Be  $expectedRows
            $sheet.Dimension.Columns                                    | Should      -Be  $columns
            $sheet.Names.Count                                          | Should      -Be ($columns + $results.Count)
            $sheet.Names[$results[0].Name]                              | Should -Not -BenullorEmpty
            $sheet.Names[$results[-1].Name]                             | Should -Not -BenullorEmpty
        }
        it "Added hyperlinks to the named ranges                                                   " {
            $sheet.cells["a1"].Hyperlink.Display                        | Should      -Match $results[0].Name
            $sheet.cells["a1"].Hyperlink.ReferenceAddress               | Should      -Match $results[0].Name
        }
    }
    Context "Adding calculated column" {
        It "Populated the cells with the right heading and formulas                                " {
            $sheet.Cells[(  $results.Count), $columns]                   | Should      -BenullorEmpty
            $sheet.Cells[(1 + $results.Count), $columns].Value             | Should      -Be "PlacesGained/Lost"
            $sheet.Cells[(2 + $results.Count), $columns].Formula           | Should      -Be "GridPosition-FinishPosition"
            $sheet.Names["PlacesGained_Lost"]                           | Should -Not -BenullorEmpty
        }
        It "Performed the calculation                                                              " {
            $placesMade = $Sheet.Cells[(2 + $results.Count), 5].value - $Sheet.Cells[(2 + $results.Count), 3].value
            $sheet.Cells[(2 + $results.Count), $columns].value             | Should -Be $placesmade
        }
        It "Applied ConditionalFormatting, including StopIfTrue, Priority                          " {
            $sheet.ConditionalFormatting[0].Address.Start.Column        | Should      -Be $columns
            $sheet.ConditionalFormatting[0].Address.End.Column          | Should      -Be $columns
            $sheet.ConditionalFormatting[0].Address.End.Row             | Should      -Be $expectedRows
            $sheet.ConditionalFormatting[0].Address.Start.Row           | Should      -Be ($results.Count + 1)
            $sheet.ConditionalFormatting[0].Icon3.Type.ToString()       | Should      -Be "Num"
            $sheet.ConditionalFormatting[0].Icon3.Value                 | Should      -Be 1
            $sheet.ConditionalFormatting[1].Priority                    | Should      -Be 1
            $sheet.ConditionalFormatting[1].StopIfTrue                  | Should      -Be $true
        }
        It "Applied ConditionalFormatting, including Reverse                                       " {
            Set-ItResult -Pending -Because "Bug in EPPLus 4.5"
            $sheet.ConditionalFormatting[3].LowValue.Color.R            | Should      -BegreaterThan 180
            $sheet.ConditionalFormatting[3].LowValue.Color.G            | Should      -BeLessThan 128
            $sheet.ConditionalFormatting[3].HighValue.Color.R           | Should      -BeLessThan 128
            $sheet.ConditionalFormatting[3].HighValue.Color.G           | Should      -BegreaterThan 180
        }
    }
    Context "Adding a table" {
        it "Created a table                                                                        " {
            $sheet.tables[0]                                            | Should -Not -BeNullOrEmpty
            $sheet.tables[0].Address.Start.Column                       | Should      -Be 1
            $sheet.tables[0].Address.End.Column                         | Should      -Be $columns
            $sheet.tables[0].Address.Start.row                          | Should      -Be ($results.Count + 1)
            $sheet.Tables[0].Address.End.Row                            | Should      -Be $expectedRows
            $sheet.Tables[0].StyleName                                  | Should      -Be "TableStyleLight4"
            $sheet.Tables[0].ShowColumnStripes                          | Should      -Be $true
            $sheet.Tables[0].ShowRowStripes                             | Should -Not -Be $true
        }
    }
    Context "Adding Pivot tables" {
        it "Added a worksheet with a pivot table grouped by date                                   " {
            $excel.Points1                                              | Should -Not -BeNullOrEmpty
            $excel.Points1.PivotTables.Count                            | Should      -Be 1
            $pt = $excel.Points1.PivotTables[0]
            $pt.RowFields.Count                                         | Should      -Be 3
            $pt.RowFields[0].name                                       | Should      -Be "Driver"
            $pt.RowFields[0].Grouping                                   | Should      -BenullorEmpty
            $pt.RowFields[1].name                                       | Should      -Be "years"
            $pt.RowFields[1].Grouping                                   | Should -Not -BenullorEmpty
            $pt.RowFields[2].name                                       | Should      -Be "date"
            $pt.RowFields[2].Grouping                                   | Should -Not -BenullorEmpty
        }
        it "Added a worksheet with a pivot table grouped by Number                                 " {
            $excel.Places1                                              | Should -Not -BeNullOrEmpty
            $excel.Places1.PivotTables.Count                            | Should      -Be 1
            $pt = $excel.Places1.PivotTables[0]
            $pt.RowFields.Count                                         | Should      -Be 2
            $pt.RowFields[0].name                                       | Should      -Be "Driver"
            $pt.RowFields[0].Grouping                                   | Should      -BenullorEmpty
            $pt.RowFields[0].SubTotalFunctions.ToString()               | Should      -Be "None"
            $pt.RowFields[1].name                                       | Should      -Be "FinishPosition"
            $pt.RowFields[1].Grouping                                   | Should -Not -BenullorEmpty
            $pt.RowFields[1].Grouping.Start                             | Should      -Be 1
            $pt.RowFields[1].Grouping.End                               | Should      -Be 25
            $pt.RowFields[1].Grouping.Interval                          | Should      -Be 3
        }
    }
    Context "Adding group date column" -Tag GroupColumnTests {
        it "Tests adding a group date column" {
            $xlFile = "TestDrive:\Results.xlsx"
            Remove-Item $xlFile -ErrorAction Ignore

            $PivotTableDefinition = New-PivotTableDefinition -Activate -PivotTableName Points `
                -PivotRows Driver -PivotColumns Date -PivotData @{Points = "SUM" } -GroupDateColumn Date -GroupDatePart Years, Months

            $excel = Import-Csv "$PSScriptRoot\First10Races.csv" |
            Select-Object Race, @{n = "Date"; e = { [datetime]::ParseExact($_.date, "dd/MM/yyyy", (Get-Culture)) } }, FinishPosition, Driver, GridPosition, Team, Points |
            Export-Excel $xlFile -AutoSize -PivotTableDefinition $PivotTableDefinition -PassThru

            $excel.Workbook.Worksheets.Count                           | Should -Be 2
            $excel.Workbook.Worksheets[1].Name                         | Should -BeExactly 'Sheet1'
            $excel.Workbook.Worksheets[2].Name                         | Should -BeExactly 'Points'
            $excel.Points.PivotTables.Count                            | Should      -Be 1
            $pt = $excel.Points.PivotTables[0]
            $pt.RowFields.Count                                        | Should      -Be 1
            $pt.RowFields[0].name                                      | Should      -Be "Driver"
            
            $pt.ColumnFields.Count                                     | Should      -Be 2

            $pt.ColumnFields[0].name                                   | Should      -Be "Years"
            $pt.ColumnFields[0].Grouping                               | Should -Not -BeNullOrEmpty
            $pt.ColumnFields[0].Grouping.GroupBy                          | Should      -Be "Years"

            $pt.ColumnFields[1].name                                   | Should      -Be "Date"
            $pt.ColumnFields[1].Grouping                               | Should -Not -BeNullOrEmpty
            $pt.ColumnFields[1].Grouping.GroupBy                       | Should      -Be "Months"

            Close-ExcelPackage $excel

            Remove-Item $xlFile -ErrorAction Ignore
        }
    }
    Context "Adding group numeric column" -Tag GroupColumnTests {
        it "Tests adding numeric group column" {
            $xlFile = "TestDrive:\Results.xlsx"
            Remove-Item $xlFile -ErrorAction Ignore

            $PivotTableDefinition = New-PivotTableDefinition -Activate -PivotTableName Places `
                -PivotRows Driver -PivotColumns FinishPosition -PivotData @{Date = "Count" } -GroupNumericColumn FinishPosition -GroupNumericMin 1 -GroupNumericMax 25 -GroupNumericInterval 3

            $excel = Import-Csv "$PSScriptRoot\First10Races.csv" |
            Select-Object  Race, @{n = "Date"; e = { [datetime]::ParseExact($_.date, "dd/MM/yyyy", (Get-Culture)) } }, FinishPosition, Driver, GridPosition, Team, Points |
            Export-Excel $xlFile -AutoSize -PivotTableDefinition $PivotTableDefinition -PassThru

            $excel.Workbook.Worksheets.Count                           | Should -Be 2
            $excel.Workbook.Worksheets[1].Name                         | Should -BeExactly 'Sheet1'
            $excel.Workbook.Worksheets[2].Name                         | Should -BeExactly 'Places'
            $excel.Places.PivotTables.Count                            | Should      -Be 1
            $pt = $excel.Places.PivotTables[0]
            $pt.RowFields.Count                                        | Should      -Be 1
            $pt.RowFields[0].name                                      | Should      -Be "Driver"
            
            $pt.ColumnFields.Count                                     | Should      -Be 1

            $pt.ColumnFields[0].name                                   | Should      -Be "FinishPosition"
            $pt.ColumnFields[0].Grouping                               | Should -Not -BeNullOrEmpty
            $pt.ColumnFields[0].Grouping.Start                         | Should      -Be 1
            $pt.ColumnFields[0].Grouping.End                           | Should      -Be 25
            $pt.ColumnFields[0].Grouping.Interval                      | Should      -Be 3
                        
            Close-ExcelPackage $excel

            Remove-Item $xlFile -ErrorAction Ignore
        }
    }
}