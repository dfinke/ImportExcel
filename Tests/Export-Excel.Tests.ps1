. $PSScriptRoot\TestHelpers.ps1

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

    Context "When piping typed data to Export-Excel using -StrictTyping switch" {

        It "Will export numeric strings that have leading zeroes as text" {
            $xlPkg = $csvData | ConvertTo-TypedObject | Export-Excel -WorkSheetname 2 $workbook -PassThru -StrictTyping
            $ws = $xlPkg.Workbook.WorkSheets[2]

            $csvData[4] | Select-Object -ExpandProperty ID | Should Be "00120"
            $ws.Cells["A6"].Value -is [string] | Should Be $true
            $ws.Cells["A6"].Value | Should Be "00120"

            $xlPkg.Save()
            $xlPkg.Dispose()
        }

        It "Will export columns with their specified types" {
            $xlPkg = $csvData | ConvertTo-TypedObject -TypeMap @{ ID = [string]; Quantity = [double] } | Export-Excel -WorkSheetname 3 $workbook -PassThru -StrictTyping
            $ws = $xlPkg.Workbook.WorkSheets[3]

            $ws.Cells["A6"].Value -is [string] | Should Be $true
            $ws.Cells["A6"].Value | Should Be "00120"

            $ws.Cells["C6"].Value -is [double] | Should Be $true
            $ws.Cells["C6"].Value | Should Be 7

            $xlPkg.Save()
            $xlPkg.Dispose()
        }

    }

    Remove-TestWorkbook
}
