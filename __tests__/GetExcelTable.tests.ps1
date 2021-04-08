if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}

Describe "Test Get Excel Table" -Tag 'Get-ExcelTable' {
    It 'Should' {
        $actual = Get-ExcelTable "$PSScriptRoot\TestXlsxFiles\MultipleSheetsAndTables.xlsx"
        
        $actual.Count | Should -Be 4    
        
        $actual[0].WorksheetName | Should -BeExactly 'Sheet1'  
        $actual[0].TableName | Should -BeExactly 'sales1'
        $actual[0].Address | Should -BeExactly 'A1:D16'
        $actual[0].Columns[0].Name | Should -BeExactly 'Region'
        $actual[0].Columns[1].Name | Should -BeExactly 'State'
        $actual[0].Columns[2].Name | Should -BeExactly 'Units'
        $actual[0].Columns[3].Name | Should -BeExactly 'Price'

        $actual[1].WorksheetName | Should -BeExactly 'Sheet1'  
        $actual[1].TableName | Should -BeExactly 'sales2'
        $actual[1].Address | Should -BeExactly 'J1:M16'
        $actual[1].Columns[0].Name | Should -BeExactly 'Region'
        $actual[1].Columns[1].Name | Should -BeExactly 'State'
        $actual[1].Columns[2].Name | Should -BeExactly 'Units'
        $actual[1].Columns[3].Name | Should -BeExactly 'Price'

        $actual[2].WorksheetName | Should -BeExactly 'SalesForJan'
        $actual[2].TableName | Should -BeExactly 'salesDataJan'
        $actual[2].Address | Should -BeExactly 'A1:D16'
        $actual[2].Columns[0].Name | Should -BeExactly 'Region'
        $actual[2].Columns[1].Name | Should -BeExactly 'State'
        $actual[2].Columns[2].Name | Should -BeExactly 'Units'
        $actual[2].Columns[3].Name | Should -BeExactly 'Price'

        $actual[3].WorksheetName | Should -BeExactly 'SalesForFeb'
        $actual[3].TableName | Should -BeExactly 'salesDataFeb'
        $actual[3].Address | Should -BeExactly 'A1:D16'
        $actual[3].Columns[0].Name | Should -BeExactly 'Region'
        $actual[3].Columns[1].Name | Should -BeExactly 'State'
        $actual[3].Columns[2].Name | Should -BeExactly 'Units'
        $actual[3].Columns[3].Name | Should -BeExactly 'Price'
    }
}