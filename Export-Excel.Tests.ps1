Import-Module $PSScriptRoot -Force

Function New-TestWorkbook {
    $TestWorkbook = Join-Path $PSScriptRoot 'Test.xlsx'
    
    Remove-Item $TestWorkbook -ErrorAction Ignore
    $TestWorkbook
}

Function Remove-TestWorkbook {
    New-TestWorkbook | Out-Null
}

Function New-TestDataCsv {
@"
ID,Product,Quantity,Price,Total
12001,Nails,37,3.99,147.63
12002,Hammer,5,12.10,60.5
12003,Saw,12,15.37,184.44
01200,Drill,20,8,160  
00120,Crowbar,7,23.48,164.36
true,Bla,7,82,12
false,Bla,7,82,12
2009-05-01 14:57:32.8,Yay,1,3,2
"@ | ConvertFrom-Csv 
}

Describe 'Export-Excel' {

    $csvData  = New-TestDataCsv
    $Workbook = New-TestWorkbook

    Context 'Importing CSV data from a here string' {
        It 'All properties are type [string]' {
            $csvData | ForEach-Object {
                $_.PSObject.Properties | ForEach-Object {
                    $_.Value -is [String] | Should Be $true
                }
            }
        }
        It 'Leading zeroes are preserved' {
            $csvData[4] | Select-Object -ExpandProperty ID | Should Be '00120'
        }
    }

    Context 'Piping CSV data to Export-Excel' {

        $Excel = $csvData | Export-Excel $Workbook -PassThru
        $Worksheet = $Excel.Workbook.Worksheets[1]

        It 'Exports numeric strings as numbers' {
            $csvData[2] | Select-Object -ExpandProperty ID | Should Be '12003'
            $Worksheet.Cells['A4'].Value -is [Double] | Should Be $true
            $Worksheet.Cells['A4'].Value | Should Be 12003
        }

        $Excel.Save()
        $Excel.Dispose()
    }

    Context 'Test PR for Appvayor' {
        it 'Test result of sum' {
            1 + 1 | Should be '2'
        }
    }

    Remove-TestWorkbook
}