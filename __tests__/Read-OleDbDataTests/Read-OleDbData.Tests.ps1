#Requires -Modules Pester
$scriptPath = $PSScriptRoot
Import-Module $scriptPath\..\..\ImportExcel.psd1 -Force
$tfp = "$scriptPath\Read-OleDbData.xlsx"
$ACEnotWorking = $false
$cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$tfp;Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'"
try {
    $Results = Read-OleDbData -ConnectionString $cs -SqlStatement "select 1"             
}
catch {
    $ACEnotWorking = $true
}
Describe "Read-OleDbData" -Tag "Read-OleDbData" {
    $PSDefaultParameterValues = @{ 'It:Skip' = $ACEnotWorking }
    Context "Basic Tests" {
        It "should be able to open spreadsheet" {
            $null = Read-OleDbData -ConnectionString $cs -SqlStatement "select 1"
            $true | Should -Be $true
        }
        It "should return PSCustomObject for single result" {
            #multiple records will come back as Object[], but not going to test for that
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement "select 1"
            $Results.GetType().Name | Should -Be 'PSCustomObject'
        }
    }
    Context "Sheet1`$A1" {
        It "Should return 1 result with a value of 1" {
            $sql = "select ROUND(F1) as [A1] from [sheet1`$A1:A1]"
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement $sql
            @($Results).length + $Results.A1 | Should -Be 2
        }
    }
    Context "Sheet2`$A1" {
        It "Should return 1 result with value of 2" {
            $sql = "select ROUND(F1) as [A1] from [sheet2`$A1:A1]"
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement $sql
            @($Results).length + $Results.A1 | Should -Be 3
        }
    }
    Context "Sheet3`$A1, Sql from file" {
        It "Should return 1 result with value of 3" {
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement (Get-Content "$scriptPath\Read-OleDbData.TestA.sql" -raw)
            @($Results).length + $Results.A1 | Should -Be 4
        }
    }
    Context "Sheets[1-7]`$A1, Sql from file" {
        It "Should return 7 result with where sum values 1-6 = value 7" {
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement (Get-Content "$scriptPath\Read-OleDbData.TestB.sql" -raw)
            $a = $Results.A1
            $a.length + ($a[0..5] | Measure-Object -sum).sum | Should -Be (7 + $a[6])
        }
    }
    Context "Sheet1`$:A1:E10, Sql from file" {
        #note, this spreadsheet doesn't have the fields populated other than A1, so it will, correctly, return only one value
        It "Should return 1 result with value of 1" {
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement (Get-Content "$scriptPath\Read-OleDbData.TestC.sql" -raw)
            @($Results).length + $Results.F1 | Should -Be 2
        }
    }
    Context "When Read-OleDbData.xlsx, select a1 from all sheets as a single record, and sql is in a file" {
        It "should return one row with 8 columns" {
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement (Get-Content "$scriptPath\Read-OleDbData.TestD.sql" -raw)
            @($Results).length + @($Results.psobject.Properties).length | Should -Be 9
        }
    }
    Context "When Read-OleDbData.xlsx, select a1 from all sheets as a single record multiple times to create a range, and sql is in a file" {
        It "should return 4 records with 5 columns" {
            $Results = Read-OleDbData -ConnectionString $cs -SqlStatement (Get-Content "$scriptPath\Read-OleDbData.TestE.sql" -raw)
            @($Results).length + @($Results[0].psobject.Properties).length | Should -Be 9
        }
    }
}