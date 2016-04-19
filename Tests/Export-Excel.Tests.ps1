#Requires -Module Pester
Set-StrictMode -Version Latest

. $PSScriptRoot\TestHelpers.ps1
. $PSScriptRoot\..\Export-Excel.ps1

Add-TypeEPPlus

Describe "Export-Excel" {

    $csvData = New-TestDataCsv
    $workbook = New-TestWorkbook

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
