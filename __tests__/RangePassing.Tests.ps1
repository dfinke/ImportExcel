#Requires -Modules Pester
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'False Positives')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingCmdletAliases', '', Justification = 'Testing for presence of alias')]
param()

describe "Consistent passing of ranges." {
    BeforeAll {
        $path = "TestDrive:\test.xlsx"
        #    if (-not (Get-command Get-Service -ErrorAction SilentlyContinue)) {
        Function Get-Service { Import-Clixml $PSScriptRoot\Mockservices.xml }
        # }
    }
    Context "Conditional Formatting" {
        it "accepts named ranges, cells['name'], worksheet + Name, worksheet + column              " {
            Remove-Item -path $path  -ErrorAction SilentlyContinue
            $excel = Get-Service | Export-Excel -Path $path -WorksheetName Services -PassThru -AutoSize -DisplayPropertySet -AutoNameRange -Title "Services on $Env:COMPUTERNAME"
            { Add-ConditionalFormatting $excel.Services.Names["Status"]  -StrikeThru -RuleType ContainsText -ConditionValue "Stopped" } | Should -Not -Throw
            $excel.Services.ConditionalFormatting.Count                                                                                | Should      -Be 1
            { Add-ConditionalFormatting $excel.Services.Cells["Name"] -Italic -RuleType ContainsText -ConditionValue "SVC" } | Should -Not -Throw
            $excel.Services.ConditionalFormatting.Count                                                                                | Should      -Be 2
            $warnvar = $null
            Add-ConditionalFormatting $excel.Services.Column(3) `
                -underline -RuleType ContainsText -ConditionValue "Windows" -WarningVariable warnvar -WarningAction SilentlyContinue
            $warnvar                                                                                                                   | Should -Not -BeNullOrEmpty
            $excel.Services.ConditionalFormatting.Count                                                                                | Should      -Be 2
            $warnvar = $null
            Add-ConditionalFormatting $excel.Services.Column(3) -Worksheet $excel.Services`
                -underline -RuleType ContainsText -ConditionValue "Windows" -WarningVariable warnvar -WarningAction SilentlyContinue
            $warnvar                                                                                                                   | Should      -BeNullOrEmpty
            $excel.Services.ConditionalFormatting.Count                                                                                | Should      -Be 3
            { Add-ConditionalFormatting "Status"  -Worksheet $excel.Services `
                    -ForeGroundColor ([System.Drawing.Color]::Green) -RuleType ContainsText -ConditionValue "Running" }                     | Should -Not -Throw
            $excel.Services.ConditionalFormatting.Count                                                                                | Should      -Be 4
            Close-ExcelPackage -NoSave $excel
        }

        it "accepts table, table.Address and worksheet + 'C:C'                                     " {
            $excel = Get-Service | Export-Excel -Path $path -WorksheetName Services -PassThru -AutoSize -DisplayPropertySet -TableName servicetable -Title "Services on $Env:COMPUTERNAME"
            { Add-ConditionalFormatting $excel.Services.Tables[0] `
                    -Italic -RuleType ContainsText -ConditionValue "Svc" } | Should -Not -Throw
            $excel.Services.ConditionalFormatting.Count                                                                                | Should      -Be 1
            { Add-ConditionalFormatting $excel.Services.Tables["ServiceTable"].Address `
                    -Bold -RuleType ContainsText -ConditionValue "windows" } | Should -Not -Throw
            $excel.Services.ConditionalFormatting.Count                                                                                | Should      -Be 2
            { Add-ConditionalFormatting -Worksheet $excel.Services -Address "a:a" `
                    -RuleType ContainsText -ConditionValue "stopped" -ForeGroundColor ([System.Drawing.Color]::Red) } | Should -Not -Throw
            $excel.Services.ConditionalFormatting.Count                                                                                | Should      -Be 3
            Close-ExcelPackage -NoSave $excel
        }
    }

    Context "Formating (Set-ExcelRange or its alias Set-Format) " {
        it "accepts Named Range, cells['Name'], cells['A1:Z9'], row, Worksheet + 'A1:Z9'" {
            $excel = Get-Service | Export-Excel -Path test2.xlsx -WorksheetName Services -PassThru -AutoSize -DisplayPropertySet -RangeName servicerange -Title "Services on $Env:COMPUTERNAME"
            { Set-format $excel.Services.Names["serviceRange"] -Bold } | Should -Not -Throw
            $excel.Services.cells["B2"].Style.Font.Bold                                                                                | Should      -Be $true
            { Set-ExcelRange -Range $excel.Services.Cells["serviceRange"] -italic:$true } | Should -Not -Throw
            $excel.Services.cells["C3"].Style.Font.Italic                                                                              | Should      -Be $true
            { Set-format $excel.Services.Row(4) -underline -Bold:$false } | Should -Not -Throw
            $excel.Services.cells["A4"].Style.Font.UnderLine                                                                           | Should      -Be $true
            $excel.Services.cells["A4"].Style.Font.Bold                                                                                | Should -Not -Be $true
            { Set-ExcelRange $excel.Services.Cells["A3:B3"] -StrikeThru } | Should -Not -Throw
            $excel.Services.cells["B3"].Style.Font.Strike                                                                              | Should      -Be $true
            { Set-ExcelRange -Worksheet $excel.Services -Range "A5:B6" -FontSize 8 } | Should -Not -Throw
            $excel.Services.cells["A5"].Style.Font.Size                                                                                | Should      -Be 8
            Close-ExcelPackage -NoSave $excel
        }

        it "Accepts Table, Table.Address , worksheet + Name, Column," {
            $excel = Get-Service | Export-Excel -Path test2.xlsx -WorksheetName Services -PassThru -AutoNameRange -AutoSize -DisplayPropertySet -TableName servicetable -Title "Services on $Env:COMPUTERNAME"
            { Set-ExcelRange $excel.Services.Tables[0] -Italic } | Should -Not -Throw
            $excel.Services.cells["C3"].Style.Font.Italic                                                                              | Should      -Be $true
            { Set-format $excel.Services.Tables["ServiceTable"].Address -Underline } | Should -Not -Throw
            $excel.Services.cells["C3"].Style.Font.UnderLine                                                                           | Should      -Be $true
            { Set-ExcelRange -Worksheet $excel.Services -Range "Name" -Bold } | Should -Not -Throw
            $excel.Services.cells["B4"].Style.Font.Bold                                                                                | Should      -Be $true
            { $excel.Services.Column(3) | Set-ExcelRange -FontColor ([System.Drawing.Color]::Red) } | Should -Not -Throw
            $excel.Services.cells["C4"].Style.Font.Color.Rgb                                                                           | Should      -Be "FFFF0000"
            Close-ExcelPackage -NoSave $excel
        }

    }

    Context "PivotTables" {
        it "Accepts Named range, .Cells['Name'], name&Worksheet, cells['A1:Z9'], worksheet&'A1:Z9' " {
            $excel = Get-Service | Export-Excel -Path $path -WorksheetName Services -PassThru -AutoSize -DisplayPropertySet -RangeName servicerange -Title "Services on $Env:COMPUTERNAME"
            $ws = $excel.Workbook.Worksheets[1] #can get a worksheet by name or index - starting at 1
            $end = $ws.Dimension.End.Address
            #can get a named ranged by name or index - starting at zero
            { Add-PivotTable -ExcelPackage $excel  -PivotTableName pt0  -SourceRange  $ws.Names[0]`
                    -PivotRows Status -PivotData Name } | Should -Not -Throw
            $excel.Workbook.Worksheets["pt0"]                                                                                          | Should -Not -BeNullOrEmpty
            { Add-PivotTable -ExcelPackage $excel  -PivotTableName pt1  -SourceRange  $ws.Names["servicerange"]`
                    -PivotRows Status -PivotData Name } | Should -Not -Throw
            $excel.Workbook.Worksheets["pt1"]                                                                                          | Should -Not -BeNullOrEmpty
            #Can specify the range for a pivot as NamedRange or Table or TableAddress or Worksheet + "A1:Z10" or worksheet + RangeName, or worksheet.cells["A1:Z10"] or worksheet.cells["RangeName"]
            { Add-PivotTable -ExcelPackage $excel  -PivotTableName pt2  -SourceRange "servicerange" -SourceWorkSheet $ws `
                    -PivotRows Status -PivotData Name } | Should -Not -Throw
            $excel.Workbook.Worksheets["pt2"]                                                                                          | Should -Not -BeNullOrEmpty
            { Add-PivotTable -ExcelPackage $excel  -PivotTableName pt3  -SourceRange  $ws.cells["servicerange"]`
                    -PivotRows Status -PivotData Name } | Should -Not -Throw
            $excel.Workbook.Worksheets["pt3"]                                                                                          | Should -Not -BeNullOrEmpty
            { Add-PivotTable -ExcelPackage $excel  -PivotTableName pt4  -SourceRange  $ws.cells["A2:$end"]`
                    -PivotRows Status -PivotData Name } | Should -Not -Throw
            $excel.Workbook.Worksheets["pt4"]                                                                                          | Should -Not -BeNullOrEmpty
            { Add-PivotTable -ExcelPackage $excel  -PivotTableName pt5  -SourceRange "A2:$end" -SourceWorkSheet $ws `
                    -PivotRows Status -PivotData Name } | Should -Not -Throw
            $excel.Workbook.Worksheets["pt5"]                                                                                           | Should -Not -BeNullOrEmpty
            Close-ExcelPackage   -NoSave $excel
        }
        it "Accepts Table, Table.Addres                                                            " {
            $excel = Get-Service | Export-Excel -Path $path -WorksheetName Services -PassThru -AutoSize -DisplayPropertySet -TableName servicetable -Title "Services on $Env:COMPUTERNAME"
            $ws = $excel.Workbook.Worksheets["Services"] #can get a worksheet by name or index - starting at 1
            #Can get a table by name or -stating at zero. Can specify the range for a pivot as or Table or TableAddress
            { Add-PivotTable -ExcelPackage $excel  -PivotTableName pt1  -SourceRange  $ws.tables["servicetable"]`
                    -PivotRows Status -PivotData Name } | Should -Not -Throw
            $excel.Workbook.Worksheets["pt1"]                                                                                          | Should -Not -BeNullOrEmpty
            { Add-PivotTable -ExcelPackage $excel  -PivotTableName pt2  -SourceRange  $ws.tables[0].Address `
                    -PivotRows Status -PivotData Name } | Should -Not -Throw
            $excel.Workbook.Worksheets["pt2"]                                                                                          | Should -Not -BeNullOrEmpty
            Close-ExcelPackage   -NoSave   $excel
        }
    }
}