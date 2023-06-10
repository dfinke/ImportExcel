if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}

Describe "Test getting the schema of an Excel file" -Tag GetExcelFileSchema {

    BeforeAll {
        $script:excelFile = "TestDrive:\test.xlsx"
        $data = ConvertFrom-Csv @"
Region,State,Units,Price
West,Texas,927,923.71
North,Tennessee,466,770.67
East,Florida,520,458.68
East,Maine,828,661.24
West,Virginia,465,053.58
North,Missouri,436,235.67
South,Kansas,214,992.47
North,North Dakota,789,640.72
South,Delaware,712,508.55
"@
        $data | Export-Excel $excelFile 
    }

    It "Test Get-ExcelFileSchema function exists" {        
        $function = Get-Command Get-ExcelFileSchema -ErrorAction SilentlyContinue
        $function | Should -Not -Be $null
    }

    It "Test Get-ExcelFileSchema returns json" {
        $actual = Get-ExcelFileSchema -Path $excelFile
        $actual | Should -Not -Be $null
        $actual | Should -BeOfType [string]        
    }

    It "Test Get-ExcelFileSchema correct json" {
        $actual = Get-ExcelFileSchema -Path $excelFile
        $actual = $actual | ConvertFrom-Json 

        $actual.ExcelFile | Should -BeExactly "test.xlsx"
        $actual.WorksheetName | Should -BeExactly "Sheet1"
        $actual.Visible | Should -Be $true
        $actual.Rows | Should -Be 10
        $actual.Columns | Should -Be 4
        $actual.Address | Should -BeExactly "A1:D10"
        $actual.Path | Should -BeExactly "TestDrive:\"
        
        $actual.PropertyNames.Count | Should -Be 4
        $actual.PropertyNames[0] | Should -BeExactly "Region"
        $actual.PropertyNames[1] | Should -BeExactly "State"
        $actual.PropertyNames[2] | Should -BeExactly "Units"
        $actual.PropertyNames[3] | Should -BeExactly "Price"
    }
}