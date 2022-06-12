[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'False Positives')]
Param()

Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force

Describe "Test reading multiple XLSX files of different row count" -Tag ReadMultipleXLSX {
    It "Should find these xlsx files" {
        Test-Path -Path $PSScriptRoot\rows05.xlsx | Should -BeTrue
        Test-Path -Path $PSScriptRoot\rows10.xlsx | Should -BeTrue
    }
    
    It "Should find two xlsx files" {
        (Get-ChildItem $PSScriptRoot\row*xlsx).Count | Should -Be 2
    }

    It "Should get 5 rows" {
        (Import-Excel $PSScriptRoot\rows05.xlsx).Count | Should -Be 5
    }

    It "Should get 10 rows" {
        (Import-Excel $PSScriptRoot\rows10.xlsx).Count | Should -Be 10
    }

    It "Should get 15 rows" {
        $actual = Get-ChildItem $PSScriptRoot\row*xlsx | Import-Excel
        
        $actual.Count | Should -Be 15
    }

    It "Should get 4 property names" {
        $actual = Get-ChildItem $PSScriptRoot\row*xlsx | Import-Excel
        
        $names = $actual[0].psobject.properties.name
        $names.Count | Should -Be 4

        $names[0] | Should -BeExactly "Region"
        $names[1] | Should -BeExactly "State"
        $names[2] | Should -BeExactly "Units"
        $names[3] | Should -BeExactly "Price"
    }
    
    It "Should have the correct data" {
        $actual = Get-ChildItem $PSScriptRoot\row*xlsx | Import-Excel
        
        # rows05.xlsx
        $actual[0].Region | Should -BeExactly "South"
        $actual[0].Price | Should -Be 181.52
        $actual[4].Region | Should -BeExactly "West"
        $actual[4].Price | Should -Be 216.56
        
        # rows10.xlsx
        $actual[5].Region | Should -BeExactly "South"
        $actual[5].Price | Should -Be 199.85
        $actual[14].Region | Should -BeExactly "East"
        $actual[14].Price | Should -Be 965.25
    }
}