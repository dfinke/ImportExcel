Describe "Test failed tests for GHA" {
    It "Should fail" {
        $true | Should -BeFalse        
    }
}