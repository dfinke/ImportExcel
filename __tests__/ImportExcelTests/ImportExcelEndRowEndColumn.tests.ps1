Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force

Describe 'Test' -Tag ImportExcelEndRowAndCols {
    BeforeAll {
        $script:xlFilename = "$PSScriptRoot\DataInDiffRowCol.xlsx"
    }

    Context 'Test reading a partial sheet' {
        It 'Should read 2 rows and first 3 columns' {
            $actual = Import-Excel $xlFilename -StartRow 5 -EndRow 7 -StartColumn 3 -EndColumn 5 

            # $actual | out-host
            $actual.Count | Should -Be 2

            $colNames = $actual[0].psobject.properties.Name
            $colNames.Count | Should -Be 3

            $colNames[0] | Should -Be 'Region'
            $colNames[1] | Should -Be 'State'
            $colNames[2] | Should -Be 'Units'
        }

        It 'Should read second 2 rows and last 2 columns' {
            $actual = Import-Excel $xlFilename -StartRow 8 -EndRow 9 -StartColumn 5 -EndColumn 6 -HeaderName 'Units', 'Price'

            # $actual | out-host
            $actual.Count | Should -Be 2

            $colNames = $actual[0].psobject.properties.Name
            $colNames.Count | Should -Be 2            

            $colNames[0] | Should -Be 'Units'
            $colNames[1] | Should -Be 'Price'
        }
    }

    Context 'Test reading multiple sheets with data in differnt rows and columns' {
        It 'Should read 2 sheets same StartRow different dimensions' {
            $xlFilename = "$PSScriptRoot\DataInDiffRowColMultipleSheets.xlsx"
            
            $actual = Import-Excel $xlFilename -StartRow 5 -WorksheetName *
            
            $actual.Keys.Count | Should -Be 2
            $actual.Contains('Sheet1') | Should -BeTrue
            $actual.Contains('Sheet2') | Should -BeTrue

            $actual['Sheet1'].Count | Should -Be 9
            $actual['Sheet2'].Count | Should -Be 12
        }
    }
}