. $PSScriptRoot\TestHelpers.ps1

. $PSScriptRoot\..\Export-Excel.ps1

Describe "Export-Excel" {

    $workbook = New-TestWorkbook
    $csvData = New-TestDataCsv

    Context "When importing CSV data from a here string" {

        It "Will create a PsCustomObject where every property value will be of type [string]" {
            $csvData | % {
                $_.PSObject.Properties | % {
                    $_.Value -is [string] | Should Be $true
                }
            }
        }
    }

    Context "When piping CSV data to Export-Excel" {

        $xlPkg = $csvData | Export-Excel $workbook -PassThru
        $ws = $xlPkg.Workbook.WorkSheets[1]

        It "Will export numeric strings as numbers" {
            $csvData[2] | Select-Object -ExpandProperty ID | Should Be "12003"
            $ws.Cells["A4"].Value -is [double] | Should Be $true
            $ws.Cells["A4"].Value | Should Be 12003
        }

        It "Will export numeric strings that have leading zeroes as numbers without the leading zeroes" {
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
