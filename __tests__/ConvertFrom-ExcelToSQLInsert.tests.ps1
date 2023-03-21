if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}

Describe "ConvertFrom-ExcelToSQLInsert" {

    BeforeAll {
        $script:xlFile = "C:\.temp\testSQL.xlsx"        # <<---------- REMOVE
        # $script:xlFile = "TestDrive:\testSQL.xlsx" 
    }

    AfterEach {
        Remove-Item $xlFile -Recurse -Force -ErrorAction Ignore
    }

    Context "When handling empty values" {

        BeforeEach {
            $([PSCustomObject]@{
                Name = "John"
                Age  = $null
            }) | Export-Excel $xlFile   
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

    Context "When escaping single quotes" {

        BeforeEach {
            $([PSCustomObject]@{
                FirstName = "John"
                LastName = "D'Angelo"
                Age  = $null
            }) | Export-Excel $xlFile
        }
    
        It "Should handle single quotes and NULLs".PadRight(90) {
            $expected = "INSERT INTO Sheet1 ('FirstName', 'LastName', 'Age') Values('John', 'D''Angelo', NULL);"
            $actual = ConvertFrom-ExcelToSQLInsert -Path $xlFile Sheet1 -SingleQuoteStyle "''" -ConvertEmptyStringsToNull
            $actual | Should -Be $expected
        }
    }
}