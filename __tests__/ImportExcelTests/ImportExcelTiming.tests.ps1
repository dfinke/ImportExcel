Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force

Describe "Tests Import-Excel Timings" -Tag Timing {
    It "Should read the 20k xlsx in -le 2100 milliseconds" {
        $timer = Measure-Command {
            $null = Import-Excel $PSScriptRoot\Rows20k.xlsx
        }

        $timer.TotalMilliseconds | Should -BeLessOrEqual 2100        
    }
}