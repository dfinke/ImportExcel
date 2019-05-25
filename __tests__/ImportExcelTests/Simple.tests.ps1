Import-Module $PSScriptRoot\..\..\ImportExcel.psd1

Describe "Tests" {
    BeforeAll {
        $data = $null
        $timer = Measure-Command {
            $data = Import-Excel $PSScriptRoot\Simple.xlsx
        }
    }

    It "Should have two items" {
        $data.count | Should be 2
    }

    It "Should have items a and b" {
        $data[0].p1 | Should be "a"
        $data[1].p1 | Should be "b"
    }

    It "Should read fast < 2100 milliseconds" {
        $timer.TotalMilliseconds | should BeLessThan 2100
    }

    It "Should read larger xlsx, 4k rows 1 col < 3000 milliseconds" {
        $timer = Measure-Command {
            $null = Import-Excel $PSScriptRoot\LargerFile.xlsx
        }

        $timer.TotalMilliseconds | should BeLessThan 3000
    }

    It "Should be able to open, read and close as seperate actions" {
        $timer = Measure-Command {
            $excel = Open-ExcelPackage $PSScriptRoot\Simple.xlsx
            $data = Import-Excel -ExcelPackage $excel
            Close-ExcelPackage -ExcelPackage $excel -NoSave}
            $timer.TotalMilliseconds | should BeLessThan 2100
            $data.count | Should be 2
            $data[0].p1 | Should be "a"
            $data[1].p1 | Should be "b"
    }
}