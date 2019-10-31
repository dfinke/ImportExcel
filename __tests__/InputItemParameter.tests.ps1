Describe "Exporting with -Inputobject; table handling, Send SQL Data and import as " {
    BeforeAll {
        $path = "TestDrive:\Results.xlsx"
        Remove-Item -Path $path -ErrorAction SilentlyContinue
        . "$PSScriptRoot\Samples\Samples.ps1"
        $results = ((Get-Process) + (Get-Process -id $PID)) | Select-Object -last  10 -Property Name, cpu, pm, handles, StartTime
        $DataTable = [System.Data.DataTable]::new('Test')
        $null = $DataTable.Columns.Add('Name')
        $null = $DataTable.Columns.Add('CPU', [double])
        $null = $DataTable.Columns.Add('PM', [Long])
        $null = $DataTable.Columns.Add('Handles', [Int])
        $null = $DataTable.Columns.Add('StartTime', [DateTime])
        foreach ($r in $results) {
            $null = $DataTable.Rows.Add($r.name, $r.CPU, $R.PM, $r.Handles, $r.StartTime)
        }
        export-excel        -Path $path -InputObject $results   -WorksheetName Sheet1 -RangeName "Whole"
        export-excel        -Path $path -InputObject $DataTable -WorksheetName Sheet2 -AutoNameRange
        Send-SQLDataToExcel -path $path -DataTable   $DataTable -WorkSheetname Sheet3  -TableName "Data"
        $DataTable.Rows.Clear()
        Send-SQLDataToExcel -path $path -DataTable   $DataTable -WorkSheetname Sheet4  -force -WarningVariable WVOne  -WarningAction SilentlyContinue
        Send-SQLDataToExcel -path $path -DataTable  ([System.Data.DataTable]::new('Test2')) -WorkSheetname Sheet5  -force -WarningVariable wvTwo -WarningAction SilentlyContinue
        $excel = Open-ExcelPackage $path
        $sheet = $excel.Sheet1
    }
    Context "Array of processes" {
        it "Put the correct rows and columns into the sheet                                        " {
            $sheet.Dimension.Rows                                       | should     be ($results.Count + 1)
            $sheet.Dimension.Columns                                    | should     be  5
            $sheet.cells["A1"].Value                                    | should     be "Name"
            $sheet.cells["E1"].Value                                    | should     be "StartTime"
            $sheet.cells["A3"].Value                                    | should     be $results[1].Name
        }
        it "Created a range for the whole sheet                                                    " {
            $sheet.Names[0].Name                                        | should     be "Whole"
            $sheet.Names[0].Start.Address                               | should     be "A1"
            $sheet.Names[0].End.row                                     | should     be ($results.Count + 1)
            $sheet.Names[0].End.Column                                  | should     be 5
        }
        it "Formatted date fields with date type                                                   " {
            $sheet.Cells["E11"].Style.Numberformat.NumFmtID             | should     be 22
        }
    }
    $sheet = $excel.Sheet2
    Context "Table of processes" {
        it "Put the correct rows and columns into the sheet                                        " {
            $sheet.Dimension.Rows                                       | should     be ($results.Count + 1)
            $sheet.Dimension.Columns                                    | should     be  5
            $sheet.cells["A1"].Value                                    | should     be "Name"
            $sheet.cells["E1"].Value                                    | should     be "StartTime"
            $sheet.cells["A3"].Value                                    | should     be $results[1].Name
        }
        it "Created named ranges for each column                                                   " {
            $sheet.Names.count                                          | should     be 5
            $sheet.Names[0].Name                                        | should     be "Name"
            $sheet.Names[1].Start.Address                               | should     be "B2"
            $sheet.Names[2].End.row                                     | should     be ($results.Count + 1)
            $sheet.Names[3].End.Column                                  | should     be 4
            $sheet.Names[4].Start.Column                                | should     be 5
        }
        it "Formatted date fields with date type                                                   " {
            $sheet.Cells["E11"].Style.Numberformat.NumFmtID             | should     be 22
        }
    }
    $sheet = $excel.Sheet3
    Context "Table of processes via Send-SQLDataToExcel" {
        it "Put the correct data rows and columns into the sheet                                   " {
            $sheet.Dimension.Rows                                       | should     be ($results.Count + 1)
            $sheet.Dimension.Columns                                    | should     be  5
            $sheet.cells["A1"].Value                                    | should     be "Name"
            $sheet.cells["E1"].Value                                    | should     be "StartTime"
            $sheet.cells["A3"].Value                                    | should     be $results[1].Name
        }
        it "Created a table                                                                        " {
            $sheet.Tables.count                                         | should     be 1
            $sheet.Tables[0].Name                                       | should     be "Data"
            $sheet.Tables[0].Columns[4].name                            | should     be "StartTime"
        }
        it "Formatted date fields with date type                                                   " {
            $sheet.Cells["E11"].Style.Numberformat.NumFmtID             | should     be 22
        }
    }
    $Sheet = $excel.Sheet4
    Context "Zero-row Data Table sent with Send-SQLDataToExcel -Force" {
        it "Raised a warning and put the correct data headers into the sheet                       " {
            $sheet.Dimension.Rows                                       | should     be  1
            $sheet.Dimension.Columns                                    | should     be  5
            $sheet.cells["A1"].Value                                    | should     be "Name"
            $sheet.cells["E1"].Value                                    | should     be "StartTime"
            $sheet.cells["A3"].Value                                    | should     beNullOrEmpty
            $wvone                                                      | should not beNullOrEmpty
        }
    }
    $Sheet = $excel.Sheet5
    Context "Zero-column Data Table handled by Send-SQLDataToExcel -Force" {
        it "Put Created a blank Sheet and raised a warning                                         " {
            $sheet.Dimension                                            | should     beNullOrEmpty
            $wvTwo                                                      | should not beNullOrEmpty
        }

    }
    Close-ExcelPackage $excel
    Context "Import As Text returns text values" {
        $x = import-excel  $path -WorksheetName sheet3 -AsText StartTime,hand* | Select-Object -last 1
        it "Had fields of type string, not date or int, where specified as ASText                  " {
            $x.Handles.GetType().Name                                   | should     be "String"
            $x.StartTime.GetType().Name                                 | should     be "String"
            $x.CPU.GetType().Name                                       | should not be "String"
        }
    }

}