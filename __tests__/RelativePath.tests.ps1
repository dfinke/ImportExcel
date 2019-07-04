Describe "Test reading relative paths" {
    It "Should read local file" {
        $actual = Import-Excel -Path ".\testRelative.xlsx"
        $actual | Should Not Be $null
        $actual.Count | Should Not Be 1
    }

    It "Should read with pwd" {
        $actual = Import-Excel -Path "$pwd\testRelative.xlsx"
        $actual | Should Not Be $null
    }

    It "Should read with PSScriptRoot" {
        $actual = Import-Excel -Path "$PSScriptRoot\testRelative.xlsx"
        $actual | Should Not Be $null
    }

    It "Should read with just a file name and resolve to cwd" {
        $actual = Import-Excel -Path "testRelative.xlsx"
        $actual | Should Not Be $null
    }
}