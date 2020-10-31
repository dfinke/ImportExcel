if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}

Describe "ConvertFrom-ExcelToSQLInsert" {
    BeforeAll {
        $script:xlFile = "TestDrive:\testSQL.xlsx"    
    }

    BeforeEach {

        $([PSCustomObject]@{
                Name = "John"
                Age  = $null
            }) | Export-Excel $xlFile
    }

    AfterAll {
        Remove-Item $xlFile -Recurse -Force -ErrorAction Ignore
    }

    It "Should be empty double single quotes".PadRight(90) {
        $expected = "INSERT INTO Sheet1 ('Name', 'Age') Values('John', '');"

        $actual = ConvertFrom-ExcelToSQLInsert -Path $xlFile Sheet1

        $actual | Should -Be $expected
    }

    It "Should have NULL".PadRight(90) {
        $expected = "INSERT INTO Sheet1 ('Name', 'Age') Values('John', NULL);"

        $actual = ConvertFrom-ExcelToSQLInsert -Path $xlFile Sheet1 -ConvertEmptyStringsToNull

        $actual | Should -Be $expected
    }
}