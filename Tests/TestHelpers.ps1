#Requires -Module Pester
#Requires -Module ImportExcel
Set-StrictMode -Version Latest

$scriptRoot = Resolve-Path $PSScriptRoot\..

. $scriptRoot\ConvertTo-TypedObject.ps1
. $scriptRoot\Export-Excel.ps1

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