if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}
Describe "Test adding trendlines to charts" {
    BeforeAll {
        $script:data = ConvertFrom-Csv @"
Region,Item,TotalSold
West,screws,60
South,lemon,48
South,apple,71
East,screwdriver,70
East,kiwi,32
West,screwdriver,1
South,melon,21
East,apple,79
South,apple,68
South,avocado,73
"@

    }

    BeforeEach {
        $xlfile = "TestDrive:\trendLine.xlsx"
        Remove-Item $xlfile -ErrorAction SilentlyContinue
    }

    It "Should add a linear trendline".PadRight(90)  {

        $cd = New-ExcelChartDefinition -XRange Region -YRange TotalSold -ChartType ColumnClustered -ChartTrendLine Linear
        $data | Export-Excel $xlfile -ExcelChartDefinition $cd -AutoNameRange

        $excel = Open-ExcelPackage -Path $xlfile
        $ws = $excel.Workbook.Worksheets["Sheet1"]

        $ws.Drawings[0].Series.TrendLines.Type | Should -Be 'Linear'

        Close-ExcelPackage $excel
    }

    It "Should add a MovingAvgerage trendline".PadRight(90)  {

        $cd = New-ExcelChartDefinition -XRange Region -YRange TotalSold -ChartType ColumnClustered -ChartTrendLine MovingAvgerage
        $data | Export-Excel $xlfile -ExcelChartDefinition $cd -AutoNameRange

        $excel = Open-ExcelPackage -Path $xlfile
        $ws = $excel.Workbook.Worksheets["Sheet1"]

        $ws.Drawings[0].Series.TrendLines.Type | Should -Be 'MovingAvgerage'

        Close-ExcelPackage $excel
    }
}