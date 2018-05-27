Import-Module $PSScriptRoot\..\..\ImportExcel.psd1
Describe "test" {

    It 'Export-Excel should be there' {
        ((Get-Command export-excel -ErrorAction SilentlyContinue) -eq $null) | Should Be $false
    }

    It "$PSScriptRoot\Simple.xlsx should exist" {
        ((ls "$PSScriptRoot\Simple.xlsx") -eq $null) | should be $false
        # Test-Path "$PSScriptRoot\Simple.xlsx" | Should Be $true
        #$PSScriptRoot\Simple.xlsx | should be "C:\projects\importexcel\UnitTests\ImportExcelTests\Simple.xlsx"
    }
}

# Import-Module $PSScriptRoot\..\..\ImportExcel.psd1
# $data = $null
# $timer = Measure-Command {
#     $data = Import-Excel $PSScriptRoot\Simple.xlsx
# }

# Describe "Tests" {
#     # BeforeAll {
#     #     $data = $null
#     #     $timer = Measure-Command {
#     #         $data = Import-Excel $PSScriptRoot\Simple.xlsx
#     #     }
#     # }

#     It "Should have two items" {
#         $data.count | Should be 2
#     }

#     It "Should have items a and b" {
#         $data[0].p1 | Should be "a"
#         $data[1].p1 | Should be "b"
#     }

#     It "Should read fast <25 milliseconds" {
#         $timer.TotalMilliseconds | should BeLessThan 25
#     }
# }