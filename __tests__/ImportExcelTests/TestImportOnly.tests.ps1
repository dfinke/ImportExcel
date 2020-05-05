Param()
Describe "Module" -Tag "TestImportOnly" {
    It "Should import without error" {
        { Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force -ErrorAction Stop } | Should -Not -Throw
    }
}
