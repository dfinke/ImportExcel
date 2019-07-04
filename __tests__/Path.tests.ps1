Describe "Test reading relative paths" {
    BeforeAll {
        $script:xlfileName = "TestR.xlsx"
        @{data = 1 } | Export-Excel "$pwd\TestR.xlsx"
    }

    AfterAll {
        Remove-Item "$pwd\$($script:xlfileName)"
    }

    It "Should read local file" {
        $actual = Import-Excel -Path ".\$($script:xlfileName)"
        $actual | Should Not Be $null
        $actual.Count | Should Be 1
    }

    It "Should read with pwd" {
        $actual = Import-Excel -Path "$pwd\$($script:xlfileName)"
        $actual | Should Not Be $null
    }

    It "Should read with just a file name and resolve to cwd" {
        $actual = Import-Excel -Path "$($script:xlfileName)"
        $actual | Should Not Be $null
    }

    It "Should fail for not found" {
        { Import-Excel -Path "ExcelFileDoesNotExist.xlsx" } | Should Throw "'ExcelFileDoesNotExist.xlsx' file not found"
    }

    It "Should fail for xls extension" {
        { Import-Excel -Path "ExcelFileDoesNotExist.xls" } | Should Throw "Import-Excel does not support reading this extension type .xls"
    }

    It "Should fail for xlsxs extension" {
        { Import-Excel -Path "ExcelFileDoesNotExist.xlsxs" } | Should Throw "Import-Excel does not support reading this extension type .xlsxs"
    }
}