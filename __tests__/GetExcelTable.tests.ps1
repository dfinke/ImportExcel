Describe "Get-ExcelTable" {

    BeforeAll {
        $script:actual = Get-ExcelTable -FullName "$PSScriptRoot\Book2.xlsx"
    }

    It "Should have 4 worksheets" {
        $actual | Should Not Be Null
        $actual.Keys.Count | Should Be 4
    }

    It "Should have these worksheet names" {
        $actual.Contains("Sheet1") | Should Be $true
        $actual.Contains("Sheet2") | Should Be $true
        $actual.Contains("Sheet3") | Should Be $true
        $actual.Contains("Cats") | Should Be $true
    }

    It "Sheet1 should have 2 table names" {
        $actual.Sheet1.Keys.Count | Should Be 2
    }

    It "Sheet1 should have these table names" {
        $actual.Sheet1.Contains("Table1") | Should Be $true
        $actual.Sheet1.Contains("Table2") | Should Be $true
    }

    It "Sheet1 Table1 should have 4 rows of data" {
        $actual.Sheet1.Table1.Count | Should Be 4
    }

    It "Sheet1 Table2 should have 4 rows of data" {
        $actual.Sheet1.Table2.Count | Should Be 4
    }

    It "Sheet2 should have 2 table names" {
        $actual.Sheet2.Keys.Count | Should Be 2
    }

    It "Sheet2 Should have these table names" {
        $actual.Sheet2.Contains("Table3") | Should Be $true
        $actual.Sheet2.Contains("Table4") | Should Be $true
    }

    It "Sheet2 Table3 should have 4 rows of data" {
        $actual.Sheet2.Table3.Count | Should Be 4
    }

    It "Sheet2 Table4 should have 4 rows of data" {
        $actual.Sheet2.Table4.Count | Should Be 4
    }

    It "Sheet3 should have 1 table name" {
        $actual.Sheet3.Keys.Count | Should Be 1
    }

    It "Sheet3 Should have this table name" {
        $actual.Sheet3.Contains("Table5") | Should Be $true
    }

    It "Sheet3 Table5 should have 4 rows of data" {
        $actual.Sheet3.Table5.Count | Should Be 4
    }

    It "Cats should have 3 table names" {
        $actual.Cats.Keys.Count | Should Be 3
    }

    It "Cats should have these table names" {
        $actual.Cats.Contains("CatFeedingTable") | Should Be $true
        $actual.Cats.Contains("CatSupplies") | Should Be $true
        $actual.Cats.Contains("FoodCosts") | Should Be $true
    }

    It "Cats CatFeedingTable should have 9 rows of data" {
        $actual.Cats.CatFeedingTable.Count | Should Be 9
    }

    It "Cats CatSupplies should have 7 rows of data" {
        $actual.Cats.CatSupplies.Count | Should Be 7
    }

    It "Cats FoodCosts should have 4 rows of data" {
        $actual.Cats.FoodCosts.Count | Should Be 4
    }

}