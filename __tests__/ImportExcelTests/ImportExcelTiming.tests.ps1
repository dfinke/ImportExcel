Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force

Describe "Tests Import-Excel Timings" -Tag Timing {
    It "Should read the 20k xlsx in -le 2100 milliseconds" {
        $timer = Measure-Command {
            $data = Import-Excel $PSScriptRoot\TimingRows20k.xlsx
        }

        $timer.TotalMilliseconds | Should -BeLessOrEqual 2100        
        $data.Count | Should -Be 19999
    }
}