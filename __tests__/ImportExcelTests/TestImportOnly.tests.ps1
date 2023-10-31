param(
    [Parameter(Mandatory)]
    [string]
    $ModulePath
)

Describe "Module" -Tag "TestImportOnly" {
    It "Should import without error" {
        { Import-Module $ModulePath -Force -ErrorAction Stop } | Should -Not -Throw
    }
}
