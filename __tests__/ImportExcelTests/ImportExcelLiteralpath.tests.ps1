Describe 'Test for using LiteralPath' -Tag LiteralPath {
    BeforeAll {
        $script:xlFilename = "$PSScriptRoot\[Test][First]File123.xlsx"
    }

    <#
    Region State        Units  Price
    ------ -----        -----  -----
    West   Texas          927 923.71
    North  Tennessee      466 770.67
    East   Florida        520 458.68
    East   Maine          828 661.24
    West   Virginia       465  53.58
    North  Missouri       436 235.67
    South  Kansas         214 992.47
    North  North Dakota   789 640.72
    South  Delaware       712 508.55
    #>

    it 'Import-Excel should read a file with wildards in them' {
 
        $actual = Import-Excel $xlFilename 
        
        $actual.Count | Should -Be 9

        $actual[0].Region | Should -BeExactly "West"
        $actual[0].State | Should -BeExactly "Texas"
        $actual[0].Units | Should -Be 927
        $actual[0].Price | Should -Be 923.71

        $actual[-1].Region | Should -BeExactly "South"
        $actual[-1].State | Should -BeExactly "Delaware"
        $actual[-1].Units | Should -Be 712
        $actual[-1].Price | Should -Be 508.55
    }

    it "Open-ExcelPackage should read a file with wildards in them" {
        $actual = Open-ExcelPackage $xlFilename

        $rows = $actual.Sheet1.Dimension.Rows
        $rows | Should -Be 10

        $actual.Sheet1.Cells[1, 1].Value | Should -BeExactly "Region"
        $actual.Sheet1.Cells[1, 2].Value | Should -BeExactly "State"
        $actual.Sheet1.Cells[1, 3].Value | Should -BeExactly "Units"
        $actual.Sheet1.Cells[1, 4].Value | Should -BeExactly "Price"

        $actual.Sheet1.Cells[2, 1].Value | Should -BeExactly "West"
        $actual.Sheet1.Cells[2, 2].Value | Should -BeExactly "Texas"
        $actual.Sheet1.Cells[2, 3].Value | Should -Be 927
        $actual.Sheet1.Cells[2, 4].Value | Should -Be 923.71

        $actual.Sheet1.Cells[$rows, 1].Value | Should -BeExactly "South"
        $actual.Sheet1.Cells[$rows, 2].Value | Should -BeExactly "Delaware"
        $actual.Sheet1.Cells[$rows, 3].Value | Should -Be 712
        $actual.Sheet1.Cells[$rows, 4].Value | Should -Be 508.55

        Close-ExcelPackage $actual -NoSave
    }
}