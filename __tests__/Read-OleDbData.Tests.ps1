Describe "Read-OleDbData" {
    BeforeAll{
        $tfp = (Get-ChildItem Read-OleDbData.xlsx).fullname # test file path
        $cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$tfp;Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'"
    }
    Context "When Read-OleDbData.xlsx and we want sheet1 a1" {
        BeforeAll{
            $sql = "select ROUND(F1) as [A1] from [sheet1`$A1:A1]"
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement $sql
        }
        It "should be PSCustomObject" {
            $Results.GetType().Name | Should -Be 'PSCustomObject'
        }
        It "should have length of 1" {
            @($Results).length | Should -Be 1
        }
        It "should be value of 1" {
            $Results.A1 | Should -Be 1
        }
    }
    Context "When Read-OleDbData.xlsx and we want sheet2 a1" {
        BeforeAll{
            $sql = "select ROUND(F1) as [A1] from [sheet2`$A1:A1]"
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement $sql
        }
        It "should be PSCustomObject" {
            $Results.GetType().Name | Should -Be 'PSCustomObject'
        }
        It "should have length of 1" {
            @($Results).length | Should -Be 1
        }
        It "should be value of 2" {
            $Results.A1 | Should -Be 2
        }
    }
    Context "When Read-OleDbData.xlsx and we want a1 on sheet3 and sql is in a file" {
        BeforeAll{
            $sql = Get-Content .\Read-OleDbData.TestA.sql -raw
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement $sql
        }
        It "should be PSCustomObject" {
            $Results.GetType().Name | Should -Be 'PSCustomObject'
        }
        It "should have length of 1" {
            @($Results).length | Should -Be 1
        }
        It "should be value of 2" {
            $Results.A1 | Should -Be 3
        }
    }
    Context "When Read-OleDbData.xlsx, we want a1 on sheets1-7, want to validate the values match properly, and sql is in a file" {
        BeforeAll{
            $sql = Get-Content .\Read-OleDbData.TestB.sql -raw
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement $sql
        }
        It "should be PSCustomObject" {
            $Results[0].GetType().Name | Should -Be 'PSCustomObject'
        }
        It "should have length of 7" {
            @($Results).length | Should -Be 7
        }
        It "should have data where sum of all initial records match the value of the last record" {
            $a = $Results.A1
            ($a[0..5] | Measure-Object -sum).sum | Should -Be $a[6]
        }
    }
    Context "When Read-OleDbData.xlsx, select range sheet1 A1:E10, and sql is in a file" {
        #note, this spreadsheet doesn't have the fields populated other than A1, so it will, correctly, return only one value
        BeforeAll{
            $sql = Get-Content .\Read-OleDbData.TestC.sql -raw
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement $sql
        }
        It "should be PSCustomObject" {
            $Results.GetType().Name | Should -Be 'PSCustomObject'
        }
        It "should have length of 1" {
            @($Results).length | Should -Be 1
        }
    }
    Context "When Read-OleDbData.xlsx, select a1 from all sheets as a single record, and sql is in a file" {
        #note, this spreadsheet doesn't have the fields populated other than A1, so it will, correctly, return only one value
        BeforeAll{
            $sql = Get-Content .\Read-OleDbData.TestD.sql -raw
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement $sql
        }
        It "should be PSCustomObject" {
            $Results.GetType().Name | Should -Be 'PSCustomObject'
        }
        It "should have length of 1" {
            @($Results).length | Should -Be 1
        }
        It "should have 8 properties" {
            @($Results.psobject.Properties).length | Should -Be 8
        }
    }
    Context "When Read-OleDbData.xlsx, select a1 from all sheets as a single record multiple times to create a range, and sql is in a file" {
        #note, this spreadsheet doesn't have the fields populated other than A1, so it will, correctly, return only one value
        BeforeAll{
            $sql = Get-Content .\Read-OleDbData.TestE.sql -raw
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement $sql
        }
        It "should be Object[]" {
            $Results.GetType().Name | Should -Be 'Object[]'
        }
        It "should have length of 4" {
            @($Results).length | Should -Be 4
        }
        It "should have 5 properties on first record" {
            @($Results[0].psobject.Properties).length | Should -Be 5
        }
    }
}
