Describe 'Read-Excel' -Tag 'Read-Excel' {
    BeforeAll {
        function Get-SheetData {
            param($SheetName)

            ConvertFrom-Csv @"
Region,State,Units,Price,SheetName
West,Texas,927,923.71,$SheetName
North,Tennessee,466,770.67,$SheetName
East,Florida,520,458.68,$SheetName
East,Maine,828,661.24,$SheetName
West,Virginia,465,053.58,$SheetName
North,Missouri,436,235.67,$SheetName
South,Kansas,214,992.47,$SheetName
North,North Dakota,789,640.72,$SheetName
South,Delaware,712,508.55,$SheetName
"@
        }
        
        $path = "TestDrive:\testReadExcel.xlsx"
        
        (Get-SheetData Sheet1) | Export-Excel -Path $path -WorksheetName Sheet1
        (Get-SheetData Sheet2) | Export-Excel -Path $path -WorksheetName Sheet2
    }

    AfterAll {
        Remove-Item $path
    }

    Context 'Test reading sheets' {
        It 'Should read all sheets' {
            # Don't specify a worksheet name
            $actual = Read-Excel $path

            $actual.Count | Should -Be 18

            $actual[0].SheetName | Should -Be "Sheet1"
            $actual[8].SheetName | Should -Be "Sheet1"
            $actual[9].SheetName | Should -Be "Sheet2"
            $actual[17].SheetName | Should -Be "Sheet2"
        }
        
        It 'Should read all sheets specified' {            
            $actual = Read-Excel $path Sheet2, Sheet1

            $actual.Count | Should -Be 18

            $actual[0].SheetName | Should -Be "Sheet2"
            $actual[8].SheetName | Should -Be "Sheet2"
            $actual[9].SheetName | Should -Be "Sheet1"
            $actual[17].SheetName | Should -Be "Sheet1"
        }

        It 'Should read one sheet' {
            $actual = Read-Excel $path sheet1

            $actual.Count | Should -Be 9

            $actual[0].SheetName | Should -Be "Sheet1"
            $actual[8].SheetName | Should -Be "Sheet1"
        }

        It 'Should fail if sheet not found' {
            { Read-Excel $path sheet1x } | Should -Throw 
        }
    }
}