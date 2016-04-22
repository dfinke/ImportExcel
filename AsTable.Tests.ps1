. $PSScriptRoot\TestHelpers.ps1

. $PSScriptRoot\..\ConvertTo-TypedObject.ps1

Describe "ConvertTo-TypedObject" {

    $csvData = New-TestDataCsv

    Context "When piping CSV data to ConvertTo-TypedObject" {

        It "Will create a new PsCustomObject where every property value will have a type dependent on its content" {
            $excelData = $csvData | ConvertTo-TypedObject
            $excelData[2] | Select-Object -ExpandProperty ID | Should Be 12003
            ($excelData[2] | Select-Object -ExpandProperty ID) -is [double] | Should Be $true
            $excelData[4] | Select-Object -ExpandProperty ID | Should Be "00120"
            ($excelData[4] | Select-Object -ExpandProperty ID) -is [string] | Should Be $true
            $excelData[5] | Select-Object -ExpandProperty ID | Should Be $true
            ($excelData[5] | Select-Object -ExpandProperty ID) -is [bool] | Should Be $true
            $excelData[6] | Select-Object -ExpandProperty ID | Should Be $false
            ($excelData[6] | Select-Object -ExpandProperty ID) -is [bool] | Should Be $true
            ($excelData[7] | Select-Object -ExpandProperty ID) -is [datetime] | Should Be $true
        }

        It "Will create a new PsCustomObject where every property value can be forced to be of a particular type per property" {
            # $excelData = $csvData | ConvertTo-TypedObject
            $excelData = $csvData | ConvertTo-TypedObject -TypeMap @{ ID=[string]; Quantity=[double] }
            $excelData | Select-Object -ExpandProperty ID | % {
                $_ -is [string] | Should Be $true
            }
            $excelData | Select-Object -ExpandProperty Quantity | % {
                $_ -is [double] | Should Be $true
            }
        }

        It "Will throw an exception when trying to force a property type that is not possible" {
            # The next section throws an exception because not every ID property could be converted to double
            { $csvData | ConvertTo-TypedObject -TypeMap @{ ID=[double]; } } | Should Throw
        }

        It "Will create a new PsCustomObject with every property value as the same type" {
            $excelData = $csvData | ConvertTo-TypedObject -GlobalType ([string])
            $excelData | % {
                $_.PSObject.Properties | % {
                    $_.Value -is [string] | Should Be $true
                }
            }
        }

        It "Will throw an exception when trying to force a global type that is not possible" {
            { $csvData | ConvertTo-TypedObject -GlobalType ([double]) } | Should Throw
        }
    }

    Context "When piping typed data to Export-Excel using -StrictTyping switch" {

        $workbook = New-TestWorkbook

        It "Will export numeric strings that have leading zeroes as text" {
            $xlPkg = $csvData | ConvertTo-TypedObject | Export-Excel $workbook -WorkSheetname test -PassThru -StrictTyping
            $ws = $xlPkg.Workbook.Worksheets[0]

            $csvData[4] | Select-Object -ExpandProperty ID | Should Be "00120"
            $ws.Cells["A6"].Value -is [string] | Should Be $true
            $ws.Cells["A6"].Value | Should Be "00120"

            $xlPkg.Save()
            $xlPkg.Dispose()
        }

        It "Will export columns with their specified types" {
            $xlPkg = $csvData | ConvertTo-TypedObject -TypeMap @{ ID = [string]; Quantity = [double] } | Export-Excel -WorkSheetname test $workbook -PassThru -StrictTyping
            $ws = $xlPkg.Workbook.WorkSheets[0]

            $ws.Cells["A6"].Value -is [string] | Should Be $true
            $ws.Cells["A6"].Value | Should Be "00120"

            $ws.Cells["C6"].Value -is [double] | Should Be $true
            $ws.Cells["C6"].Value | Should Be 7

            $xlPkg.Save()
            $xlPkg.Dispose()
        }
    }
}
