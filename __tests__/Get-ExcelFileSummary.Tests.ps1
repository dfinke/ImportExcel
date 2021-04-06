#Requires -Modules Pester
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'False Positives')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidAssignmentToAutomaticVariable', '', Justification = 'Only executes on versions without the automatic variable')]
param()

if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}

Describe 'All tests for Get-ExcelFileSummary' -Tag "Get-ExcelFileSummary" {
    Context "Test Get-ExcelFileSummary" {
        It "Tests summary on TestData2.xlsx" {
            $actual = Get-ExcelFileSummary "$PSScriptRoot\ImportExcelTests\TestData1.xlsx"

            $actual.ExcelFile | Should -BeExactly 'TestData1.xlsx'
            $actual.WorksheetName | Should -BeExactly 'Sheet1'
            $actual.Rows | Should -Be 3
            $actual.Columns | Should -Be 2
            $actual.Address | Should -BeExactly 'A1:B3'
            $actual.Path | Should -Not -BeNullOrEmpty
        }

        It "Tests summary on xlsx with multiple sheets" {

            $actual = Get-ExcelFileSummary "$PSScriptRoot\ImportExcelTests\MultipleSheets.xlsx"

            $actual[0].ExcelFile | Should -BeExactly 'MultipleSheets.xlsx'
            $actual[0].WorksheetName | Should -BeExactly 'Sheet1'
            $actual[0].Rows | Should -Be 1
            $actual[0].Columns | Should -Be 4
            $actual[0].Address | Should -BeExactly 'A1:D1'
            $actual[0].Path | Should -Not -BeNullOrEmpty

            $actual[1].ExcelFile | Should -BeExactly 'MultipleSheets.xlsx'
            $actual[1].WorksheetName | Should -BeExactly 'Sheet2'
            $actual[1].Rows | Should -Be 2
            $actual[1].Columns | Should -Be 2
            $actual[1].Address | Should -BeExactly 'A1:B2'
            $actual[1].Path | Should -Not -BeNullOrEmpty
        }

    }
}
