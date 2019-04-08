Describe "Exporting with -Inputobject" {
    BeforeAll {
        $path = "$env:TEMP\Results.xlsx"
        Remove-Item -Path $path -ErrorAction SilentlyContinue
        #Read race results, and group by race name : export 1 row to get headers, leaving enough rows aboce to put in a link for each race
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
        it "Put the correct rows and columns into the sheet                                        " {
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
}