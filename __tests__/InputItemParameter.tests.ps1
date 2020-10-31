# [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','',Justification='False Positives')]
# Param()
# Describe "Exporting with -Inputobject, table handling, Send-SQL-Data. Checking Import -asText" {
#     BeforeAll {
#         $path  = "TestDrive:\Results.xlsx"
#         $path2 = "TestDrive:\Results2.xlsx"
#         Remove-Item -Path $path,$path2 -ErrorAction SilentlyContinue
#         if (Test-path "$PSScriptRoot\Samples\Samples.ps1") {. "$PSScriptRoot\Samples\Samples.ps1"}
#         $results = ((Get-Process) + (Get-Process -id $PID)) | Select-Object -last  10 -Property Name, cpu, pm, handles, StartTime
#         $DataTable = [System.Data.DataTable]::new('Test')
#         $null = $DataTable.Columns.Add('Name')
#         $null = $DataTable.Columns.Add('CPU', [double])
#         $null = $DataTable.Columns.Add('PM', [Long])
#         $null = $DataTable.Columns.Add('Handles', [Int])
#         $null = $DataTable.Columns.Add('StartTime', [DateTime])
#         Send-SQLDataToExcel -path $path -DataTable   $DataTable -WorkSheetname Sheet4  -force -TableName "Data" -WarningVariable WVOne  -WarningAction SilentlyContinue
#         Send-SQLDataToExcel -path $path -DataTable  ([System.Data.DataTable]::new('Test2')) -WorkSheetname Sheet5  -force -WarningVariable wvTwo -WarningAction SilentlyContinue
#         foreach ($r in $results) {
#             $null = $DataTable.Rows.Add($r.name, $r.CPU, $R.PM, $r.Handles, $r.StartTime)
#         }
#         $NowPkg   =  Export-Excel -InputObject $DataTable -PassThru
#         $NowPath1 = $NowPkg.File.FullName
#         Close-ExcelPackage $NowPkg
#         $NowPkg   = Export-Excel -InputObject $DataTable -PassThru -table:$false
#         $NowPath2 = $NowPkg.File.FullName
#         Close-ExcelPackage $NowPkg
#         Export-Excel        -Path $path -InputObject $results   -WorksheetName Sheet1 -RangeName "Whole"
#         Export-Excel        -Path $path -InputObject $DataTable -WorksheetName Sheet2 -AutoNameRange
#         Send-SQLDataToExcel -path $path -DataTable   $DataTable -WorkSheetname Sheet3 -TableName "Data" -WarningVariable WVThree -WarningAction SilentlyContinue

#         Send-SQLDataToExcel -Path $path2 -DataTable $DataTable -WorksheetName Sheet1 -Append
#         Send-SQLDataToExcel -Path $path2 -DataTable $DataTable -WorksheetName Sheet1 -Append

#         Send-SQLDataToExcel -Path $path2 -DataTable $DataTable -WorksheetName Sheet2 -Append -TableName "FirstLot" -TableStyle light7
#         Send-SQLDataToExcel -Path $path2 -DataTable $DataTable -WorksheetName Sheet2 -Append

#         Send-SQLDataToExcel -Path $path2 -DataTable $DataTable -WorksheetName Sheet3 -Append
#         Send-SQLDataToExcel -Path $path2 -DataTable $DataTable -WorksheetName Sheet3 -Append -TableName "SecondLot"

#         Send-SQLDataToExcel -Path $path2 -DataTable $DataTable -WorksheetName Sheet4 -Append
#         Send-SQLDataToExcel -Path $path2 -DataTable $DataTable -WorksheetName Sheet4 -Append -TableStyle  Dark5


#         $excel = Open-ExcelPackage $path
#         $sheet = $excel.Sheet1
#     }
#     Context "Array of processes" {
#         it "Put the correct rows and columns into the sheet                                        " {
#             $sheet.Dimension.Rows                                       | Should      -Be ($results.Count + 1)
#             $sheet.Dimension.Columns                                    | Should      -Be  5
#             $sheet.cells["A1"].Value                                    | Should      -Be "Name"
#             $sheet.cells["E1"].Value                                    | Should      -Be "StartTime"
#             $sheet.cells["A3"].Value                                    | Should      -Be $results[1].Name
#         }
#         it "Created a range for the whole sheet                                                    " {
#             $sheet.Names[0].Name                                        | Should      -Be "Whole"
#             $sheet.Names[0].Start.Address                               | Should      -Be "A1"
#             $sheet.Names[0].End.row                                     | Should      -Be ($results.Count + 1)
#             $sheet.Names[0].End.Column                                  | Should      -Be 5
#         }
#         it "Formatted date fields with date type                                                   " {
#             $sheet.Cells["E11"].Style.Numberformat.NumFmtID             | Should      -Be 22
#         }
#     }
#     $sheet = $excel.Sheet2
#     Context "Table of processes" {
#         it "Put the correct rows and columns into the sheet                                        " {
#             $sheet.Dimension.Rows                                       | Should      -Be ($results.Count + 1)
#             $sheet.Dimension.Columns                                    | Should      -Be  5
#             $sheet.cells["A1"].Value                                    | Should      -Be "Name"
#             $sheet.cells["E1"].Value                                    | Should      -Be "StartTime"
#             $sheet.cells["A3"].Value                                    | Should      -Be $results[1].Name
#         }
#         it "Created named ranges for each column                                                   " {
#             $sheet.Names.count                                          | Should      -Be 5
#             $sheet.Names[0].Name                                        | Should      -Be "Name"
#             $sheet.Names[1].Start.Address                               | Should      -Be "B2"
#             $sheet.Names[2].End.row                                     | Should      -Be ($results.Count + 1)
#             $sheet.Names[3].End.Column                                  | Should      -Be 4
#             $sheet.Names[4].Start.Column                                | Should      -Be 5
#         }
#         it "Formatted date fields with date type                                                   " {
#             $sheet.Cells["E11"].Style.Numberformat.NumFmtID             | Should      -Be 22
#         }
#     }

#     Context "'Now' Mode behavior" {
#         $NowPkg = Open-ExcelPackage $NowPath1
#         $sheet = $NowPkg.Sheet1
#         it "Formatted data as a table by default                                                   " {
#             $sheet.Tables.Count                                         | Should      -Be  1
#         }
#         Close-ExcelPackage -NoSave $NowPkg
#         Remove-Item $NowPath1
#         $NowPkg = Open-ExcelPackage $NowPath2
#         $sheet = $NowPkg.Sheet1
#         it "Did not data as a table when table:`$false was used                                     " {
#             $sheet.Tables.Count                                         | Should      -Be  0
#         }
#         Close-ExcelPackage -NoSave $NowPkg
#         Remove-Item $NowPath2
#     }
#     $sheet = $excel.Sheet3
#     Context "Table of processes via Send-SQLDataToExcel" {
#         it "Put the correct data rows and columns into the sheet                                   " {
#             $sheet.Dimension.Rows                                       | Should      -Be ($results.Count + 1)
#             $sheet.Dimension.Columns                                    | Should      -Be  5
#             $sheet.cells["A1"].Value                                    | Should      -Be "Name"
#             $sheet.cells["E1"].Value                                    | Should      -Be "StartTime"
#             $sheet.cells["A3"].Value                                    | Should      -Be $results[1].Name
#         }
#         it "Created a table                                                                        " {
#             $sheet.Tables.count                                         | Should      -Be 1
#             $sheet.Tables[0].Columns[4].name                            | Should      -Be "StartTime"
#         }
#         it "Formatted date fields with date type                                                   " {
#             $sheet.Cells["E11"].Style.Numberformat.NumFmtID             | Should      -Be 22
#         }
#         it "Handled two data tables with the same name                                             " {
#             $sheet.Tables[0].Name                                       | Should      -Be "Data_"
#             $wvThree[0]                                                 | Should      -Match "is not unique"
#         }
#     }
#     $Sheet = $excel.Sheet4
#     Context "Zero-row Data Table sent with Send-SQLDataToExcel -Force" {
#         it "Raised a warning and put the correct data headers into the sheet                       " {
#             $sheet.Dimension.Rows                                       | Should      -Be  1
#             $sheet.Dimension.Columns                                    | Should      -Be  5
#             $sheet.cells["A1"].Value                                    | Should      -Be "Name"
#             $sheet.cells["E1"].Value                                    | Should      -Be "StartTime"
#             $sheet.cells["A3"].Value                                    | Should      -BeNullOrEmpty
#             $wvone[0]                                                   | Should      -Match "Zero"
#         }
#         it "Applied table formatting                                                               " {
#             $sheet.Tables.Count                                         | Should      -Be  1
#             $sheet.Tables[0].Name                                       | Should      -Be "Data"
#         }


#     }
#     $Sheet = $excel.Sheet5
#     Context "Zero-column Data Table handled by Send-SQLDataToExcel -Force" {
#         it "Created a blank Sheet and raised a warning                                             " {
#             $sheet.Dimension                                            | Should      -BeNullOrEmpty
#             $wvTwo                                                      | Should -Not -BeNullOrEmpty
#         }

#     }
#     Close-ExcelPackage $excel
#     $excel = Open-ExcelPackage $path2
#     Context "Send-SQLDataToExcel -append works correctly" {
#         it "Works without table settings                                                           " {
#             $excel.sheet1.Dimension.Address                             | Should      -Be "A1:E21"
#             $excel.sheet1.cells[1,1].value                              | Should      -Be "Name"
#             $excel.sheet1.cells[12,1].value                             | Should      -Be $excel.sheet1.cells[2,1].value
#             $excel.sheet1.Tables.count                                  | Should      -Be  0
#         }
#         it "Extends an existing table when appending                                               " {
#             $excel.sheet2.Dimension.Address                             | Should      -Be "A1:E21"
#             $excel.sheet2.cells[1,2].value                              | Should      -Be "CPU"
#             $excel.sheet2.cells[13,2].value                             | Should      -Be $excel.sheet2.cells[3,2].value
#             $excel.sheet2.Tables.count                                  | Should      -Be  1
#             $excel.sheet2.Tables[0].name                                | Should      -Be "FirstLot"
#             $excel.sheet2.Tables[0].StyleName                           | Should      -Be "TableStyleLight7"
#         }
#         it "Creates a new table by name when appending                                             " {
#             $excel.sheet3.cells[1,3].value                              | Should      -Be "PM"
#             $excel.sheet3.cells[14,3].value                             | Should      -Be $excel.sheet3.cells[4,3].value
#             $excel.sheet3.Tables.count                                  | Should      -Be  1
#             $excel.sheet3.Tables[0].name                                | Should      -Be "SecondLot"
#             $excel.sheet3.Tables[0].StyleName                           | Should      -Be "TableStyleMedium6"
#         }
#         it "Creates a new table by style when appending                                            " {
#             $excel.sheet4.cells[1,4].value                              | Should      -Be "Handles"
#             $excel.sheet4.cells[15,4].value                             | Should      -Be $excel.sheet4.cells[5,4].value
#             $excel.sheet4.Tables.count                                  | Should      -Be  1
#             $excel.sheet4.Tables[0].name                                | Should      -Be "Table1"
#             $excel.sheet4.Tables[0].StyleName                           | Should      -Be "TableStyleDark5"
#         }
#     }

#     Close-ExcelPackage $excel
#     Context "Import As Text returns text values" {
#         $x = Import-excel  $path -WorksheetName sheet3 -AsText StartTime,hand* | Select-Object -last 1
#         it "Had fields of type string, not date or int, where specified as ASText                  " {
#             $x.Handles.GetType().Name                                   | Should      -Be "String"
#             $x.StartTime.GetType().Name                                 | Should      -Be "String"
#             $x.CPU.GetType().Name                                       | Should -Not -Be "String"
#         }
#     }

# }