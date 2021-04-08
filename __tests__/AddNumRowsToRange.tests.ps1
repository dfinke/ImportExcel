if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}

Describe "Test Add Num To Rows" -Tag 'Add-NumRowsToRange' {
    It 'Should add 1 to the range' {
        $actual = Add-NumRowsToRange -range 'A1:B1'

        $actual | Should -BeExactly 'A1:B2'
    }

    It 'Should add 5 to the range' {
        $actual = Add-NumRowsToRange -range 'B3:C3' -numRowsToAdd 5

        $actual | Should -BeExactly 'B3:C8'
    }

    # It 'Should add 1 to the range' {
    #     $actual = Add-NumRowsToRange -range 'A1'

    #     $actual | Should -BeExactly 'B3:C8'
    # }
}