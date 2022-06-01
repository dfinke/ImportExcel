Describe "Test reading relative paths" {
    BeforeAll {
        $script:xlfileName = "TestR.xlsx"
        @{data = 1 } | Export-Excel (Join-Path $PWD  "TestR.xlsx")
    }

    AfterAll {
        Remove-Item (Join-Path $PWD  "$($script:xlfileName)")
    }

    It "Should read local file".PadRight(90) {
        $actual = Import-Excel -Path ".\$($script:xlfileName)"
        $actual | Should -Not -Be $null
        $actual.Count | Should -Be 1
    }

    It "Should read with pwd".PadRight(90) {
        $actual = Import-Excel -Path (Join-Path $PWD  "$($script:xlfileName)")
        $actual | Should -Not -Be $null
    }

    It "Should read with just a file name and resolve to cwd".PadRight(90) {
        $actual = Import-Excel -Path "$($script:xlfileName)"
        $actual | Should -Not -Be $null
    }

    It "Should fail for not found".PadRight(90) {
        { Import-Excel -Path "ExcelFileDoesNotExist.xlsx" } | Should -Throw "Failed importing the Excel workbook 'ExcelFileDoesNotExist.xlsx' with worksheet '': File not found"
    }

    It "Should fail for xls extension".PadRight(90) {
        { Import-Excel -Path "ExcelFileDoesNotExist.xls" } | Should -Throw "Failed importing the Excel workbook 'ExcelFileDoesNotExist.xls' with worksheet '': Extension type '.xls' not supported"
    }

    It "Should fail for xlsxs extension".PadRight(90) {
        { Import-Excel -Path "ExcelFileDoesNotExist.xlsxs" } | Should -Throw "Failed importing the Excel workbook 'ExcelFileDoesNotExist.xlsxs' with worksheet '': Extension type '.xlsxs' not supported"
    }

    It "Should produce only one error on failure".PadRight(90) {
        $error.clear()
        { Import-Excel -Path "ExcelFileDoesNotExist.xls" } | Should -Throw
        $error | Should -HaveCount 1
    }
}