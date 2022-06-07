#Requires -Modules Pester
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'False Positives')]
param()

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
            $actual[99].Month | Should -BeExactly "April"
        }

        It 'Should read two sheets' {
            $actual = Import-Excel $xlFilename march, june

            $actual.keys.Count | Should -Be 2
            $actual["March"].Count | Should -Be 100
            $actual["June"].Count | Should -Be 100
        }

        It 'Should read all the sheets' {
            $actual = Import-Excel $xlFilename *

            $actual.keys.Count | Should -Be 12

            $actual["January"].Count | Should -Be 100
            $actual["February"].Count | Should -Be 100
            $actual["March"].Count | Should -Be 100
            $actual["April"].Count | Should -Be 100
            $actual["May"].Count | Should -Be 100
            $actual["June"].Count | Should -Be 100
            $actual["July"].Count | Should -Be 100
            $actual["August"].Count | Should -Be 100
            $actual["September"].Count | Should -Be 100
            $actual["October"].Count | Should -Be 100
            $actual["November"].Count | Should -Be 100
            $actual["December"].Count | Should -Be 100
        }

        It 'Should throw if it cannot find the sheet' {
            { Import-Excel $xlFilename april, june, notthere } | Should -Throw
        }

        It 'Should return an array not a dictionary' {
            $actual = Import-Excel $xlFilename april, june -Raw
            
            $actual.Count | Should -Be 200
            $group = $actual | Group-Object month -NoElement

            $group.Count | Should -Be 2
            $group[0].Name | Should -BeExactly 'April'
            $group[1].Name | Should -BeExactly 'June'
        }

        It "Should read multiple sheets with diff number of rows correctly" {
            $xlFilename = "$PSScriptRoot\construction.xlsx"

            $actual = Import-Excel $xlFilename 2015, 2016
            $actual.keys.Count | Should -Be 2

            $actual["2015"].Count | Should -Be 12
            $actual["2016"].Count | Should -Be 1
        }

        It "Should read multiple sheets with diff number of rows correctly and flatten it" {
            $xlFilename = "$PSScriptRoot\construction.xlsx"

            $actual = Import-Excel $xlFilename 2015, 2016 -Raw

            $actual.Count | Should -Be 13
        }

    }
}