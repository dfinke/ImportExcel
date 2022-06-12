Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force

Describe 'Test' -Tag ImportExcelEndRowAndCols {
    BeforeAll {
        $script:xlFilename = "$PSScriptRoot\DataInDiffRowCol.xlsx"
    }

    Context 'Test reading a patial sheet' {
        It 'Should read 2 rows and first 3 columns' {
            $actual = Import-Excel $xlFilename -StartRow 5 -EndRow 7 -StartColumn 3 -EndColumn 5 

            $actual | out-host
            $actual.Count | Should -Be 2

            $colNames = $actual[0].psobject.properties.Name
            $colNames.Count | Should -Be 3

            $colNames[0] | Should -Be 'Region'
            $colNames[1] | Should -Be 'State'
            $colNames[2] | Should -Be 'Units'
        }

        It 'Should read second 2 rows and last 2 columns' {
            $actual = Import-Excel $xlFilename -StartRow 8 -EndRow 9 -StartColumn 5 -EndColumn 6 -HeaderName 'Units', 'Price'

            $actual | out-host
            $actual.Count | Should -Be 2

            $colNames = $actual[0].psobject.properties.Name
            $colNames.Count | Should -Be 2            

            $colNames[0] | Should -Be 'Units'
            $colNames[1] | Should -Be 'Price'
        }
    }
}