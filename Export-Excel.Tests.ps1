#Requires -Module Pester
#Requires -Module ImportExcel
Set-StrictMode -Version Latest

. $PSScriptRoot\Export-Excel.ps1

function New-TestWorkbook {
    $testWorkbook = Join-Path $PSScriptRoot test.xlsx
    if (Test-Path $testWorkbook) {
        rm $testWorkbook -Force
    }
    $testWorkbook
}

function Remove-TestWorkbook {
    New-TestWorkbook | Out-Null
}

function New-TestDataCsv {
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

function Get-PathEPPlusDll {
    Resolve-Path $PSScriptRoot\..\EPPlus.dll
}

Describe "Export-Excel" {

    $csvData = New-TestDataCsv
    $workbook = New-TestWorkbook

    Context "Importing CSV data from a here string" {
        It "All properties are type [string]" {
            $csvData | % {
                $_.PSObject.Properties | % {
                    $_.Value -is [string] | Should Be $true
                }
            }
        }
        It "Leading zeroes are preserved" {
            $csvData[4] | Select-Object -ExpandProperty ID | Should Be "00120"
        }
    }

    Context "Piping CSV data to Export-Excel" {

        $xlPkg = $csvData | Export-Excel $workbook -PassThru
        $ws = $xlPkg.Workbook.WorkSheets[1]

        It "Exports numeric strings as numbers" {
            $csvData[2] | Select-Object -ExpandProperty ID | Should Be "12003"
            $ws.Cells["A4"].Value -is [double] | Should Be $true
            $ws.Cells["A4"].Value | Should Be 12003
        }

        It "Exports numeric strings that have leading zeroes as numbers without the leading zeroes" {
            $csvData[4] | Select-Object -ExpandProperty ID | Should Be "00120"
            $ws.Cells["A6"].Value -is [double] | Should Be $true
            $ws.Cells["A6"].Value | Should Be 120
            $ws.Cells["A6"].Value | Should Not Be "00120"
        }

        $xlPkg.Save()
        $xlPkg.Dispose()
    }

    Remove-TestWorkbook
}
