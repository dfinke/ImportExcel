$scriptPath      = Split-Path -Path $MyInvocation.MyCommand.path -Parent
$dataPath        = Join-Path  -Path $scriptPath -ChildPath "First10Races.csv"

Describe "Creating small named ranges with hyperlinks" {
    BeforeAll {
        $path    = "$env:TEMP\Results.xlsx"
        Remove-Item -Path $path -ErrorAction SilentlyContinue
        #Read race results, and group by race name : export 1 row to get headers, leaving enough rows aboce to put in a link for each race
        $results = Import-Csv -Path $dataPath | Group-Object -Property RACE
        $topRow  = $lastDataRow = 1 + $results.Count
        $excel   = $results[0].Group[0] | Export-Excel -Path $path -StartRow $TopRow  -BoldTopRow -PassThru

        #export each group (race) below the last one, without headers, and create a range for each using the group name (Race)
        foreach ($r in $results) {
            $excel        = $R.Group | Export-Excel -ExcelPackage $excel -NoHeader -StartRow ($lastDataRow +1) -RangeName $R.Name -PassThru -AutoSize
            $lastDataRow += $R.Group.Count
        }
        $worksheet = $excel.Workbook.Worksheets[1]
        $columns   = $worksheet.Dimension.Columns

        1..$columns | ForEach-Object {Add-ExcelName -Range $worksheet.cells[$topRow,$_,$lastDataRow,$_]}                                                                        #Test Add-Excel Name on its own (outside Export-Excel)

        $scwarnVar = $null
        Set-ExcelColumn -Worksheet $worksheet -StartRow $topRow -Heading "PlacesGained/Lost" `
                                        -Value "=GridPosition-FinishPosition" -AutoNameRange -WarningVariable scWarnVar -WarningAction SilentlyContinue                 #Test as many set column options as possible.
        $columns ++

        #create a table which covers all the data. And define a pivot table which uses the same address range.
        $table = Add-ExcelTable -PassThru  -Range  $worksheet.cells[$topRow,1,$lastDataRow,$columns]  -TableName "AllResults" -TableStyle Light4 `
                                    -ShowHeader -ShowFilter -ShowColumnStripes -ShowRowStripes:$false -ShowFirstColumn:$false -ShowLastColumn:$false -ShowTotal:$false   #Test Add-ExcelTable outside export-Excel with as many options as possible.
        $pt = New-PivotTableDefinition -PivotTableName Analysis -SourceWorkSheet   $worksheet -SourceRange $table.address.address -PivotRows Driver -PivotData @{Points="SUM"} -PivotTotals None

        $cf = Add-ConditionalFormatting -Address  $worksheet.cells[$topRow,$columns,$lastDataRow,$columns] -ThreeIconsSet Arrows  -Passthru                               #Test using cells[r1,c1,r2,c2]
        $cf.Icon2.Type = $cf.Icon3.Type = "Num"
        $cf.Icon2.Value = 0
        $cf.Icon3.Value = 1
        Add-ConditionalFormatting -Address $worksheet.cells["FinishPosition"] -RuleType Equal    -ConditionValue 1 -ForeGroundColor Purple -Bold -Priority 1 -StopIfTrue   #Test Priority and stopIfTrue and using range name
        Add-ConditionalFormatting -Address $worksheet.Cells["GridPosition"]   -RuleType ThreeColorScale -Reverse                                                           #Test Reverse
        $ct = New-ConditionalText -Text "Ferrari"
        $ct2 =  New-ConditionalText -Range $worksheet.Names["FinishPosition"].Address -ConditionalType LessThanOrEqual -Text 3 -ConditionalText Red -Background White      #Test new-conditionalText in shortest and longest forms.
        #Create links for each group name (race) and Export them so they start at Cell A1; create a pivot table with definition just created, save the file and open in Excel
        $results | ForEach-Object {(New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList "Sheet1!$($_.Name)" , "$($_.name) GP")} |                                     #Test Exporting Hyperlinks with display property.
            Export-Excel -ExcelPackage $excel -AutoSize -PivotTableDefinition $pt -Calculate   -ConditionalFormat $ct,$ct2                                                 #Test conditional text rules in conditional format (orignally icon sets only )

        $excel = Open-ExcelPackage $path
        $sheet = $excel.Workbook.Worksheets[1]
        $m = $results | measure -sum -Property count
        $expectedRows = 1 + $m.count + $m.sum
    }
    Context "Creating hyperlinks" {
        it "Put the data into the sheet and created the expected named ranges                      " {
            $sheet.Dimension.Rows                                       | should     be  $expectedRows
            $sheet.Dimension.Columns                                    | should     be  $columns
            $sheet.Names.Count                                          | should     be ($columns + $results.Count)
            $sheet.Names[$results[0].Name]                              | should not benullorEmpty
            $sheet.Names[$results[-1].Name]                             | should not benullorEmpty
        }
        it "Added hyperlinks to the named ranges                                                   " {
            $sheet.cells["a1"].Hyperlink.Display                        | should     match $results[0].Name
            $sheet.cells["a1"].Hyperlink.ReferenceAddress               | should     match $results[0].Name
        }
    }
    Context "Adding calculated column" {
        It "Populated the cells with the right heading and formulas                                " {
            $sheet.Cells[(  $results.Count),$columns]                   | Should     benullorEmpty
            $sheet.Cells[(1+$results.Count),$columns].Value             | Should     be "PlacesGained/Lost"
            $sheet.Cells[(2+$results.Count),$columns].Formula           | should     be "GridPosition-FinishPosition"
            $sheet.Names["PlacesGained_Lost"]                           | should not benullorEmpty
        }
        It "Performed the calculation                                                              " {
            $placesMade = $Sheet.Cells[(2+$results.Count),5].value - $Sheet.Cells[(2+$results.Count),3].value
            $sheet.Cells[(2+$results.Count),$columns].value             | Should be $placesmade
        }
        It "Applied ConditionalFormatting, including StopIfTrue, Priority and Reverse              " {
            $sheet.ConditionalFormatting[0].Address.Start.Column        | should     be $columns
            $sheet.ConditionalFormatting[0].Address.End.Column          | should     be $columns
            $sheet.ConditionalFormatting[0].Address.End.Row             | should     be $expectedRows
            $sheet.ConditionalFormatting[0].Address.Start.Row           | should     be ($results.Count + 1)
            $sheet.ConditionalFormatting[0].Icon3.Type.ToString()       | Should     be "Num"
            $sheet.ConditionalFormatting[0].Icon3.Value                 | Should     be 1
            $sheet.ConditionalFormatting[1].Priority                    | Should     be 1
            $sheet.ConditionalFormatting[1].StopIfTrue                  | Should     be $true
            $sheet.ConditionalFormatting[3].LowValue.Color.R            | Should     begreaterThan 180
            $sheet.ConditionalFormatting[3].LowValue.Color.G            | Should     beLessThan 128
            $sheet.ConditionalFormatting[3].HighValue.Color.R           | Should     beLessThan 128
            $sheet.ConditionalFormatting[3].HighValue.Color.G           | Should     begreaterThan 180
        }
    }
    Context "Adding a table" {
        it "Created a table                                                                        " {
            $sheet.tables[0]                                            | Should not beNullOrEmpty
            $sheet.tables[0].Address.Start.Column                       | should     be 1
            $sheet.tables[0].Address.End.Column                         | should     be $columns
            $sheet.tables[0].Address.Start.row                          | should     be ($results.Count + 1)
            $sheet.Tables[0].Address.End.Row                            | should     be $expectedRows
            $sheet.Tables[0].StyleName                                  | should     be "TableStyleLight4"
            $sheet.Tables[0].ShowColumnStripes                          | should     be $true
            $sheet.Tables[0].ShowRowStripes                             | should not be $true
        }
    }
}