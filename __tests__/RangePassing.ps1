$path = "$env:temp\test.xlsx"
describe "Consistent passing of ranges." {
    Context "Conditional Formatting"  {
        Remove-Item -path $path  -ErrorAction SilentlyContinue
        $excel = Get-Service | Export-Excel -Path $path -WorksheetName Services -PassThru -AutoSize -DisplayPropertySet -AutoNameRange -Title "Services on $Env:COMPUTERNAME"
        it "accepts named ranges, cells['name'], worksheet + Name, worksheet + column              " {
            {Add-ConditionalFormatting $excel.Services.Names["Status"]  -StrikeThru -RuleType ContainsText -ConditionValue "Stopped" } | Should not throw
            $excel.Services.ConditionalFormatting.Count                                                                                | Should     be 1
            {Add-ConditionalFormatting $excel.Services.Cells["Name"] -Italic -RuleType ContainsText -ConditionValue "SVC"            } | Should not throw
            $excel.Services.ConditionalFormatting.Count                                                                                | Should     be 2
            $warnvar = $null
            Add-ConditionalFormatting $excel.Services.Column(3) `
                -underline -RuleType ContainsText -ConditionValue "Windows" -WarningVariable warnvar -WarningAction SilentlyContinue
            $warnvar                                                                                                                   | should not beNullOrEmpty
            $excel.Services.ConditionalFormatting.Count                                                                                | Should     be 2
            $warnvar = $null
            Add-ConditionalFormatting $excel.Services.Column(3) -WorkSheet $excel.Services`
            -underline -RuleType ContainsText -ConditionValue "Windows" -WarningVariable warnvar -WarningAction SilentlyContinue
            $warnvar                                                                                                                   | should     beNullOrEmpty
            $excel.Services.ConditionalFormatting.Count                                                                                | Should     be 3
            {Add-ConditionalFormatting "Status"  -WorkSheet $excel.Services `
                -ForeGroundColor ([System.Drawing.Color]::Green) -RuleType ContainsText -ConditionValue "Running"}                     | Should not throw
            $excel.Services.ConditionalFormatting.Count                                                                                | Should     be 4
        }
        Close-ExcelPackage -NoSave $excel
        $excel = Get-Service | Export-Excel -Path $path -WorksheetName Services -PassThru -AutoSize -DisplayPropertySet -TableName servicetable -Title "Services on $Env:COMPUTERNAME"
        it "accepts table, table.Address and worksheet + 'C:C'                                     " {
            {Add-ConditionalFormatting $excel.Services.Tables[0] `
                -Italic -RuleType ContainsText -ConditionValue "Svc"                                                                 } | Should not throw
            $excel.Services.ConditionalFormatting.Count                                                                                | Should     be 1
            {Add-ConditionalFormatting $excel.Services.Tables["ServiceTable"].Address `
                -Bold -RuleType ContainsText -ConditionValue "windows"                                                               } | Should not throw
            $excel.Services.ConditionalFormatting.Count                                                                                | Should     be 2
            {Add-ConditionalFormatting -WorkSheet $excel.Services -Address "a:a" `
                -RuleType ContainsText -ConditionValue "stopped" -ForeGroundColor ([System.Drawing.Color]::Red)                      } | Should not throw
            $excel.Services.ConditionalFormatting.Count                                                                                | Should     be 3
        }
        Close-ExcelPackage -NoSave $excel
    }

    Context "Formating (Set-ExcelRange or its alias set-Format) " {
        it "accepts Named Range, cells['Name'], cells['A1:Z9'], row, Worksheet + 'A1:Z9'" {
            $excel = Get-Service | Export-Excel -Path test2.xlsx -WorksheetName Services -PassThru -AutoSize -DisplayPropertySet -RangeName servicerange -Title "Services on $Env:COMPUTERNAME"
            {set-format $excel.Services.Names["serviceRange"] -Bold                                                                  } | Should Not Throw
            $excel.Services.cells["B2"].Style.Font.Bold                                                                                | Should     be $true
            {Set-ExcelRange -Range $excel.Services.Cells["serviceRange"] -italic:$true                                               } | Should not throw
            $excel.Services.cells["C3"].Style.Font.Italic                                                                              | Should     be $true
            {set-format $excel.Services.Row(4) -underline -Bold:$false                                                               } | Should not throw
            $excel.Services.cells["A4"].Style.Font.UnderLine                                                                           | Should     be $true
            $excel.Services.cells["A4"].Style.Font.Bold                                                                                | Should not be $true
            {Set-ExcelRange $excel.Services.Cells["A3:B3"] -StrikeThru                                                               } | Should not throw
            $excel.Services.cells["B3"].Style.Font.Strike                                                                              | Should     be $true
            {Set-ExcelRange -WorkSheet $excel.Services -Range "A5:B6" -FontSize 8                                                    } | Should not throw
            $excel.Services.cells["A5"].Style.Font.Size                                                                                | Should     be 8
        }
        Close-ExcelPackage -NoSave $excel
        it "Accepts Table, Table.Address , worksheet + Name, Column," {
            $excel = Get-Service | Export-Excel -Path test2.xlsx -WorksheetName Services -PassThru -AutoNameRange -AutoSize -DisplayPropertySet -TableName servicetable -Title "Services on $Env:COMPUTERNAME"
            {set-ExcelRange $excel.Services.Tables[0] -Italic                                                                        } | Should not throw
            $excel.Services.cells["C3"].Style.Font.Italic                                                                              | Should     be $true
            {set-format $excel.Services.Tables["ServiceTable"].Address -Underline                                                    } | Should not throw
            $excel.Services.cells["C3"].Style.Font.UnderLine                                                                           | Should     be $true
            {Set-ExcelRange -WorkSheet $excel.Services -Range "Name" -Bold                                                           } | Should not throw
            $excel.Services.cells["B4"].Style.Font.Bold                                                                                | Should     be $true
           {$excel.Services.Column(3) | Set-ExcelRange -FontColor ([System.Drawing.Color]::Red)                                      } | Should not throw
            $excel.Services.cells["C4"].Style.Font.Color.Rgb                                                                           | Should     be "FFFF0000"
        }
        Close-ExcelPackage -NoSave $excel
    }

    Context "PivotTables" {
        it "Accepts Named range, .Cells['Name'], name&Worksheet, cells['A1:Z9'], worksheet&'A1:Z9' "{
            $excel = Get-Service | Export-Excel -Path $path -WorksheetName Services -PassThru -AutoSize -DisplayPropertySet -RangeName servicerange -Title "Services on $Env:COMPUTERNAME"
            $ws    = $excel.Workbook.Worksheets[1] #can get a worksheet by name or index - starting at 1
            $end   = $ws.Dimension.End.Address
            #can get a named ranged by name or index - starting at zero
            {Add-PivotTable -ExcelPackage $excel  -PivotTableName pt0  -SourceRange  $ws.Names[0]`
                -PivotRows Status -PivotData Name                                                                                    } | Should not throw
            $excel.Workbook.Worksheets["pt0"]                                                                                          | Should not beNullOrEmpty
            {Add-PivotTable -ExcelPackage $excel  -PivotTableName pt1  -SourceRange  $ws.Names["servicerange"]`
                    -PivotRows Status -PivotData Name                                                                                } | Should not throw
            $excel.Workbook.Worksheets["pt1"]                                                                                          | Should not beNullOrEmpty
            #Can specify the range for a pivot as NamedRange or Table or TableAddress or Worksheet + "A1:Z10" or worksheet + RangeName, or worksheet.cells["A1:Z10"] or worksheet.cells["RangeName"]
            {Add-PivotTable -ExcelPackage $excel  -PivotTableName pt2  -SourceRange "servicerange" -SourceWorkSheet $ws `
                    -PivotRows Status -PivotData Name                                                                                } | Should not throw
            $excel.Workbook.Worksheets["pt2"]                                                                                          | Should not beNullOrEmpty
            {Add-PivotTable -ExcelPackage $excel  -PivotTableName pt3  -SourceRange  $ws.cells["servicerange"]`
                    -PivotRows Status -PivotData Name                                                                                } | Should not throw
            $excel.Workbook.Worksheets["pt3"]                                                                                          | Should not beNullOrEmpty
            {Add-PivotTable -ExcelPackage $excel  -PivotTableName pt4  -SourceRange  $ws.cells["A2:$end"]`
                    -PivotRows Status -PivotData Name                                                                                } | Should not throw
            $excel.Workbook.Worksheets["pt4"]                                                                                          | Should not beNullOrEmpty
            {Add-PivotTable -ExcelPackage $excel  -PivotTableName pt5  -SourceRange "A2:$end" -SourceWorkSheet $ws `
                    -PivotRows Status -PivotData Name                                                                                 } | Should not throw
            $excel.Workbook.Worksheets["pt5"]                                                                                           | Should not beNullOrEmpty
             Close-ExcelPackage   -NoSave $excel
        }
        it "Accepts Table, Table.Addres                                                            " {
            $excel = Get-Service | Export-Excel -Path $path -WorksheetName Services -PassThru -AutoSize -DisplayPropertySet -TableName servicetable -Title "Services on $Env:COMPUTERNAME"
            $ws    = $excel.Workbook.Worksheets["Services"] #can get a worksheet by name or index - starting at 1
            #Can get a table by name or -stating at zero. Can specify the range for a pivot as or Table or TableAddress
            {Add-PivotTable -ExcelPackage $excel  -PivotTableName pt1  -SourceRange  $ws.tables["servicetable"]`
                    -PivotRows Status -PivotData Name                                                                               } | Should not throw
            $excel.Workbook.Worksheets["pt1"]                                                                                          | Should not beNullOrEmpty
            {Add-PivotTable -ExcelPackage $excel  -PivotTableName pt2  -SourceRange  $ws.tables[0].Address `
                    -PivotRows Status -PivotData Name                                                                                } | Should not throw
            $excel.Workbook.Worksheets["pt2"]                                                                                          | Should not beNullOrEmpty
            Close-ExcelPackage   -NoSave   $excel
        }



    }
}