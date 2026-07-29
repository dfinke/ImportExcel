Import-Module $PSScriptRoot\..\ImportExcel.psd1 -Force

Describe "Re-exporting tables and autofilters to an existing worksheet - Issue #1725" {
    # Exporting into an existing worksheet used to stack a second table over the first (when no
    # -TableName was given) or combine a table with a leftover worksheet autofilter. Excel treats
    # both as a corrupt file ("We found a problem with some content...") and strips the table
    # style while repairing. Each scenario below must yield exactly one table and no clashing
    # worksheet autofilter.
    BeforeAll {
        $data = ConvertFrom-Csv -InputObject @"
Mail,List
foo.bar,Nein
"@
        $data2 = ConvertFrom-Csv -InputObject @"
Mail,List
foo.bar,Nein
baz.qux,Ja
"@
        $wsName = "Geteilte Postfächer"
    }
    Context "The same -TableStyle export run twice against one file" {
        BeforeAll {
            $path = "TestDrive:\rerun.xlsx"
            $data | Export-Excel -Path $path -WorksheetName $wsName -AutoFilter -TableStyle Light2
            $data | Export-Excel -Path $path -WorksheetName $wsName -AutoFilter -TableStyle Light2
            $excel = Open-ExcelPackage -Path $path
            $ws = $excel.Workbook.Worksheets[$wsName]
        }
        AfterAll { Close-ExcelPackage -ExcelPackage $excel -NoSave }
        it "Left a single table with the requested style rather than two stacked tables            " {
            $ws.Tables.Count                                            | Should      -Be 1
            $ws.Tables[0].Address.Address                               | Should      -Be 'A1:B2'
            $ws.Tables[0].StyleName                                     | Should      -Be 'TableStyleLight2'
        }
        it "Did not add a worksheet autofilter on top of the table                                 " {
            $ws.AutoFilterAddress                                       | Should      -BeNullOrEmpty
        }
    }
    Context "A -TableStyle export over data that grew since the previous export" {
        BeforeAll {
            $path = "TestDrive:\grow.xlsx"
            $data  | Export-Excel -Path $path -WorksheetName $wsName -TableStyle Light2
            $data2 | Export-Excel -Path $path -WorksheetName $wsName -TableStyle Light2
            $excel = Open-ExcelPackage -Path $path
            $ws = $excel.Workbook.Worksheets[$wsName]
        }
        AfterAll { Close-ExcelPackage -ExcelPackage $excel -NoSave }
        it "Stretched the existing table over the new rows instead of adding a second table       " {
            $ws.Tables.Count                                            | Should      -Be 1
            $ws.Tables[0].Address.Address                               | Should      -Be 'A1:B3'
        }
    }
    Context "A -TableStyle export into a sheet left with an autofilter by an earlier export" {
        BeforeAll {
            $path = "TestDrive:\afthentable.xlsx"
            $data | Export-Excel -Path $path -WorksheetName $wsName -AutoFilter
            $data | Export-Excel -Path $path -WorksheetName $wsName -AutoFilter -TableStyle Light2
            $excel = Open-ExcelPackage -Path $path
            $ws = $excel.Workbook.Worksheets[$wsName]
        }
        AfterAll { Close-ExcelPackage -ExcelPackage $excel -NoSave }
        it "Removed the leftover autofilter when it created the table                              " {
            $ws.AutoFilterAddress                                       | Should      -BeNullOrEmpty
            $ws.Tables.Count                                            | Should      -Be 1
            $ws.Tables[0].StyleName                                     | Should      -Be 'TableStyleLight2'
        }
    }
    Context "An -AutoFilter export into a sheet left with a table by an earlier export" {
        BeforeAll {
            $path = "TestDrive:\tablethenaf.xlsx"
            $data | Export-Excel -Path $path -WorksheetName $wsName -TableStyle Light2
            $data | Export-Excel -Path $path -WorksheetName $wsName -AutoFilter -WarningVariable afWarning -WarningAction SilentlyContinue
            $excel = Open-ExcelPackage -Path $path
            $ws = $excel.Workbook.Worksheets[$wsName]
        }
        AfterAll { Close-ExcelPackage -ExcelPackage $excel -NoSave }
        it "Warned, and left filtering to the table instead of adding an autofilter over it       " {
            $afWarning                                                  | Should -Not -BeNullOrEmpty
            $ws.AutoFilterAddress                                       | Should      -BeNullOrEmpty
            $ws.Tables.Count                                            | Should      -Be 1
        }
    }
    Context "An unnamed -TableStyle export over a previously named table" {
        BeforeAll {
            $path = "TestDrive:\namedthenplain.xlsx"
            $data  | Export-Excel -Path $path -WorksheetName $wsName -TableName MailListe -TableStyle Light2
            $data2 | Export-Excel -Path $path -WorksheetName $wsName -TableStyle Light9
            $excel = Open-ExcelPackage -Path $path
            $ws = $excel.Workbook.Worksheets[$wsName]
        }
        AfterAll { Close-ExcelPackage -ExcelPackage $excel -NoSave }
        it "Re-used the named table, stretching it and applying the new style                     " {
            $ws.Tables.Count                                            | Should      -Be 1
            $ws.Tables[0].Name                                          | Should      -Be 'MailListe'
            $ws.Tables[0].Address.Address                               | Should      -Be 'A1:B3'
            $ws.Tables[0].StyleName                                     | Should      -Be 'TableStyleLight9'
        }
    }
    Context "A named -TableName export over a previously unnamed table" {
        BeforeAll {
            $path = "TestDrive:\plainthennamed.xlsx"
            $data | Export-Excel -Path $path -WorksheetName $wsName -TableStyle Light2
            $data | Export-Excel -Path $path -WorksheetName $wsName -TableName MailListe -TableStyle Light9
            $excel = Open-ExcelPackage -Path $path
            $ws = $excel.Workbook.Worksheets[$wsName]
        }
        AfterAll { Close-ExcelPackage -ExcelPackage $excel -NoSave }
        it "Took over the existing table, renaming it, rather than doubling up                    " {
            $ws.Tables.Count                                            | Should      -Be 1
            $ws.Tables[0].TableXml.table.name                           | Should      -Be 'MailListe'
            $ws.Tables[0].StyleName                                     | Should      -Be 'TableStyleLight9'
        }
    }
    Context "A -TableStyle export over data whose columns changed since the previous export" {
        BeforeAll {
            $wide = ConvertFrom-Csv -InputObject @"
Mail,List,Extra
foo.bar,Nein,x
"@
            $path = "TestDrive:\widen.xlsx"
            $data | Export-Excel -Path $path -WorksheetName $wsName -TableStyle Light2
            $wide | Export-Excel -Path $path -WorksheetName $wsName -TableStyle Light2
            $excel = Open-ExcelPackage -Path $path
            $ws = $excel.Workbook.Worksheets[$wsName]
        }
        AfterAll { Close-ExcelPackage -ExcelPackage $excel -NoSave }
        it "Rebuilt the table's column definitions to match the new width                         " {
            $ws.Tables.Count                                            | Should      -Be 1
            $ws.Tables[0].TableXml.table.ref                            | Should      -Be 'A1:C2'
            $ws.Tables[0].TableXml.table.autoFilter.ref                 | Should      -Be 'A1:C2'
            $ws.Tables[0].TableXml.table.tableColumns.count             | Should      -Be 3
            $ws.Tables[0].TableXml.table.tableColumns.tableColumn.name -join ',' | Should -Be 'Mail,List,Extra'
        }
    }
    Context "A widening export which also turns the totals row on" {
        BeforeAll {
            $d2 = ConvertFrom-Csv -InputObject "Name,Amount`nAlpha,1`nBeta,2"
            $w2 = ConvertFrom-Csv -InputObject "Name,Amount,Extra`nAlpha,1,x`nBeta,2,y"
            $path = "TestDrive:\widentotals.xlsx"
            $d2 | Export-Excel -Path $path -WorksheetName S -TableName WideTot -TableStyle Light2
            $w2 | Export-Excel -Path $path -WorksheetName S -TableName WideTot -TableStyle Light2 -TableTotalSettings @{Extra='Count';Amount='Sum'}
            $excel = Open-ExcelPackage -Path $path
            $tableXml = $excel.Workbook.Worksheets['S'].Tables[0].TableXml.table
        }
        AfterAll { Close-ExcelPackage -ExcelPackage $excel -NoSave }
        it "Kept the range, filter range and column definitions consistent                        " {
            $tableXml.ref                                               | Should      -Be 'A1:C4'
            $tableXml.totalsRowCount                                    | Should      -Be '1'
            $tableXml.autoFilter.ref                                    | Should      -Be 'A1:C3'
            $tableXml.tableColumns.count                                | Should      -Be 3
        }
    }
    Context "A re-export over a table which shows a totals row" {
        BeforeAll {
            $path = "TestDrive:\totals.xlsx"
            $pkg = Open-ExcelPackage -Path $path -Create
            $wsT = $pkg.Workbook.Worksheets.Add('S')
            $wsT.Cells['A1'].Value = 'Name'; $wsT.Cells['B1'].Value = 'Amount'
            $wsT.Cells['A2'].Value = 'Alpha'; $wsT.Cells['B2'].Value = 1
            Add-ExcelTable -Range $wsT.Cells['A1:B2'] -TableName TotTbl -ShowTotal
            Close-ExcelPackage $pkg
            $d3 = ConvertFrom-Csv -InputObject "Name,Amount`nAlpha,1`nBeta,2`nGamma,3"
            $d3 | Export-Excel -Path $path -WorksheetName S -TableName TotTbl -TableStyle Light2
            $excel = Open-ExcelPackage -Path $path
            $tableXml = $excel.Workbook.Worksheets['S'].Tables[0].TableXml.table
        }
        AfterAll { Close-ExcelPackage -ExcelPackage $excel -NoSave }
        it "Kept a row for the totals below the data and excluded it from the filter range        " {
            $tableXml.ref                                               | Should      -Be 'A1:B5'
            $tableXml.totalsRowCount                                    | Should      -Be '1'
            $tableXml.autoFilter.ref                                    | Should      -Be 'A1:B4'
        }
    }
    Context "Renaming a table in a package which stays open between exports" {
        BeforeAll {
            $d2 = ConvertFrom-Csv -InputObject "Name,Amount`nAlpha,1`nBeta,2"
            $d3 = ConvertFrom-Csv -InputObject "Name,Amount`nAlpha,1`nBeta,2`nGamma,3"
            $renameWarnings = @()
            $pkg = $d2 | Export-Excel -Path "TestDrive:\rename.xlsx" -WorksheetName S -TableStyle Light1 -PassThru
            $pkg = $d2 | Export-Excel -ExcelPackage $pkg -WorksheetName S -TableName Renamed -TableStyle Light1 -PassThru -WarningVariable +renameWarnings
            $pkg = $d3 | Export-Excel -ExcelPackage $pkg -WorksheetName S -TableName Renamed -TableStyle Light9 -PassThru -WarningVariable +renameWarnings
            $ws = $pkg.Workbook.Worksheets['S']
        }
        AfterAll { Close-ExcelPackage -ExcelPackage $pkg -NoSave }
        it "Found the renamed table again on the next export instead of warning                   " {
            $renameWarnings                                             | Should      -BeNullOrEmpty
            $ws.Tables.Count                                            | Should      -Be 1
            $ws.Tables[0].TableXml.table.name                           | Should      -Be 'Renamed'
            $ws.Tables[0].TableXml.table.ref                            | Should      -Be 'A1:B4'
            $ws.Tables[0].StyleName                                     | Should      -Be 'TableStyleLight9'
        }
    }
}
