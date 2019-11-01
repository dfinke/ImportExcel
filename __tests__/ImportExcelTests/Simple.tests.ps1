if (-not (get-command Import-Excel -ErrorAction SilentlyContinue)) {
     Import-Module $PSScriptRoot\..\..\ImportExcel.psd1
}
Describe "Tests" {
    BeforeAll {
        $data = $null
        $timer = Measure-Command {
            $data = Import-Excel $PSScriptRoot\Simple.xlsx
        }
    }

    It "Should have two items".PadRight(90) {
        $data.count | Should be 2
    }

    It "Should have items a and b".PadRight(90) {
        $data[0].p1 | Should be "a"
        $data[1].p1 | Should be "b"
    }

    It "Should read fast < 2100 milliseconds".PadRight(90) {
        $timer.TotalMilliseconds | should BeLessThan 2100
    }

    It "Should read larger xlsx, 4k rows 1 col < 3000 milliseconds".PadRight(90) {
        $timer = Measure-Command {
            $null = Import-Excel $PSScriptRoot\LargerFile.xlsx
        }

        $timer.TotalMilliseconds | should BeLessThan 3000
    }

    It "Should be able to open, read and close as seperate actions".PadRight(90) {
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