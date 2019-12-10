if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}
$xlFile = "TestDrive:\testSQL.xlsx"

Describe "ConvertFrom-ExcelToSQLInsert" {

    BeforeEach {

      @([pscustomobject]@{name = "John"; age = $null},
      [pscustomobject]@{name = "Matthew"; age = 20},
      [pscustomobject]@{name = "Mark"; age = 30},
      [pscustomobject]@{name = "O'Rourke"; age = 40},
      [pscustomobject]@{name = "Luke"; age = 50})| Export-Excel $xlFile

    }

    AfterAll {
        Remove-Item $xlFile -Recurse -Force -ErrorAction Ignore
    }

    It "Should be empty double single quotes".PadRight(90)  {
        $expected="INSERT INTO Sheet1 ('Name', 'Age') Values('John', '');"

        $actual = ConvertFrom-ExcelToSQLInsert -Path $xlFile Sheet1

        $actual[0] | should be $expected
    }

     It "Should have NULL".PadRight(90)  {
        $expected="INSERT INTO Sheet1 ('Name', 'Age') Values('John', NULL);"

        $actual = ConvertFrom-ExcelToSQLInsert -Path $xlFile Sheet1 -ConvertEmptyStringsToNull

        $actual[0] | should be $expected
    }

    It "Should support the StartRows argument".PadRight(90)  {
      $expected = "INSERT INTO Sheet1 ('O''Rourke', '40') Values('Luke', '50');"

      $actual = ConvertFrom-ExcelToSQLInsert -Path $xlFile -Tablename Sheet1 -StartRow 5

      $actual | should be $expected
    }

    It "Should put brackets around names for MSSQLSyntax".PadRight(90)  {
      $expected = "INSERT INTO Sheet1 ([O'Rourke], [40]) Values('Luke', '50');"

      $actual = ConvertFrom-ExcelToSQLInsert -Path $xlFile -Tablename Sheet1 -StartRow 5 -UseMSSQLSyntax

      $actual | should be $expected
    }

    It "Should escape apostrophes".PadRight(90)  {
          $expected="INSERT INTO Sheet1 ('name', 'age') Values('O''Rourke', '40');"

          $actual = ConvertFrom-ExcelToSQLInsert -Path $xlFile Sheet1

          $actual[3] | should be $expected
    }
}