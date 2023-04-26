Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 -Force

Describe "Tests Import-Excel Timings" -Tag Timing {
    It "Should read the 20k xlsx in -le 2100 milliseconds" {
        $timer = Measure-Command {
            $data = Import-Excel $PSScriptRoot\TimingRows20k.xlsx
        }

        $timer.TotalMilliseconds | Should -BeLessOrEqual 2100
        $data.Count | Should -Be 19999
    }
    It "Should read the 20k xlsx in -le 5000 milliseconds all sheets" {
        $timer = Measure-Command {
            $data = Import-Excel $PSScriptRoot\TimingRows20k.xlsx *
        }

        $data.Count | Should -be 3
        $timer.TotalMilliseconds | Should -BeLessOrEqual 5000        
        $data[0].Count | Should -Be 19999
        $data[1].Count | Should -Be 10000
        $data[2].Count | Should -Be 13824
    }
}