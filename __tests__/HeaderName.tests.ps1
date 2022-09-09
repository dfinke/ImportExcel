if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}

Describe "Test HeaderName parameter" -Tag HeaderName {
    It "Should add data as usual" {
        $xlfile = "TestDrive:\headername.xlsx"
        Remove-Item $xlfile -ErrorAction SilentlyContinue

        1..10 | Export-Excel -Path $xlfile

        { Test-Path $xlfile } | Should -BeTrue

        $excel = Open-ExcelPackage $xlfile
        
        1..10 | ForEach-Object {
            $excel.Sheet1.Cells[$_, 1].Text | Should -BeExactly $_
        }

        Close-ExcelPackage $excel

        Remove-Item $xlfile -ErrorAction SilentlyContinue
    }

    It "Should add data and the first cell should have the header name" {
        $xlfile = "TestDrive:\headername.xlsx"
        Remove-Item $xlfile -ErrorAction SilentlyContinue

        1..10 | Export-Excel -Path $xlfile -HeaderName MyNum

        { Test-Path $xlfile } | Should -BeTrue

        $excel = Open-ExcelPackage $xlfile

        $excel.Sheet1.Cells[1, 1].Text | Should -BeExactly "MyNum"

        1..10 | ForEach-Object {
            $excel.Sheet1.Cells[($_ + 1), 1].Text | Should -BeExactly $_
        }

        Close-ExcelPackage $excel

        Remove-Item $xlfile -ErrorAction SilentlyContinue
    }

    It "Should ignore the header name" {
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

        $xlfile = "TestDrive:\headername.xlsx"
        Remove-Item $xlfile -ErrorAction SilentlyContinue

        $data | Export-Excel -Path $xlfile -HeaderName MyNum

        { Test-Path $xlfile } | Should -BeTrue

        $excel = Open-ExcelPackage $xlfile

        $excel.Sheet1.Cells[1, 1].Text | Should -BeExactly "Region"
        $excel.Sheet1.Cells[1, 2].Text | Should -BeExactly "State"
        $excel.Sheet1.Cells[1, 3].Text | Should -BeExactly "Units"
        $excel.Sheet1.Cells[1, 4].Text | Should -BeExactly "Price"

        $excel.Sheet1.Cells[2, 1].Text | Should -BeExactly "West"
        $excel.Sheet1.Cells[2, 2].Text | Should -BeExactly "Texas"
        $excel.Sheet1.Cells[2, 3].Text | Should -Be 927
        $excel.Sheet1.Cells[2, 4].Text | Should -Be 923.71

        $excel.Sheet1.Cells[10, 1].Text | Should -BeExactly "South"
        $excel.Sheet1.Cells[10, 2].Text | Should -BeExactly "Delaware"
        $excel.Sheet1.Cells[10, 3].Text | Should -Be 712
        $excel.Sheet1.Cells[10, 4].Text | Should -Be 508.55        

        Close-ExcelPackage $excel

        Remove-Item $xlfile -ErrorAction SilentlyContinue
    }
}