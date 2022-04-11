Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force

Describe 'Different ways to import sheets' -Tag ImportExcelReadSheets {
    BeforeAll {
        $xlFilename = "$PSScriptRoot\yearlySales.xlsx"
    }

    Context 'Test reading sheets' {
        It 'Should read one sheet' {
            $actual = Import-Excel $xlFilename

            $actual.Count | Should -Be 100
            $actual[0].Month | Should -BeExactly "April"
        }

        It 'Should read two sheets' {
            $actual = Import-Excel $xlFilename march, june

            $actual.Count | Should -Be 200
            $actual[0].Month | Should -BeExactly "March"
            $actual[100].Month | Should -BeExactly "June"
        }

        It 'Should read all the sheets' {
            $actual = Import-Excel $xlFilename *

            $actual.Count | Should -Be 1200
            
            $actual[0].Month | Should -BeExactly "April"
            $actual[100].Month | Should -BeExactly "August"
            $actual[200].Month | Should -BeExactly "December"
            $actual[300].Month | Should -BeExactly "February"
            $actual[400].Month | Should -BeExactly "January"
            $actual[500].Month | Should -BeExactly "July"
            $actual[600].Month | Should -BeExactly "June"
            $actual[700].Month | Should -BeExactly "March"
            $actual[800].Month | Should -BeExactly "May"
            $actual[900].Month | Should -BeExactly "November"
            $actual[1000].Month | Should -BeExactly "October"
            $actual[1100].Month | Should -BeExactly "September"
        }

        It 'Should throw if it cannot find the sheet' {
            { Import-Excel $xlFilename april, june, notthere } | Should -Throw
        }
    }
}