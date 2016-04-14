#Requires -Module Pester
#Requires -Module ImportExcel
Set-StrictMode -Version Latest

# Import-Module ImportExcel -Force

function New-TestWorkbook {
    $testWorkbook = Join-Path $PSScriptRoot test.xlsx
    if (Test-Path $testWorkbook) {
        rm $testWorkbook -Force
    }
    $testWorkbook
}

function Remove-TestWorkbook {
    Write-Host "Removing test workbook."
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
"@ | ConvertFrom-Csv 
}

Describe "ExportSimple" {
    Context "When piping CSV data to Export-Excel" {
        It "Exports numeric strings as numbers and not text" {
            $workbook = New-TestWorkbook
            $data = New-TestDataCsv

            $xlPkg = $data | Export-Excel $workbook -PassThru
            $ws = $xlPkg.Workbook.WorkSheets[1]

            $data[2] | Select-Object -ExpandProperty ID | Should Be "12003"
            $data[2] | Select-Object -ExpandProperty ID | % { $_.GetType().Name } | Should Be "string"
            $ws.Cells["A4"].Value | Should Be 12003

            $data[4] | Select-Object -ExpandProperty ID | Should Be "00120"
            $data[4] | Select-Object -ExpandProperty ID | % { $_.GetType().Name } | Should Be "string"
            $ws.Cells["A6"].Value | Should Be 120
            $ws.Cells["A6"].Value | Should Not Be "00120"

            $xlPkg.Save()
            $xlPkg.Dispose()
        }
    }
}

Remove-TestWorkbook
