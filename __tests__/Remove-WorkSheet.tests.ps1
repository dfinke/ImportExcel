#Requires -Modules Pester
Import-Module $PSScriptRoot\..\ImportExcel.psd1 -Force

Describe "Remove Worksheet" {
    Context "Remove a worksheet output" {
        BeforeEach {
            # Create three sheets
            $data = ConvertFrom-Csv @"
Name,Age
Jane,10
John,20
"@
            $xlFile1 = "$env:TEMP\RemoveWorsheet1.xlsx"
            Remove-Item $xlFile1 -ErrorAction SilentlyContinue

            $data | Export-Excel -Path $xlFile1 -WorksheetName Target1
            $data | Export-Excel -Path $xlFile1 -WorksheetName Target2
            $data | Export-Excel -Path $xlFile1 -WorksheetName Target3
            $data | Export-Excel -Path $xlFile1 -WorksheetName Sheet1

            $xlFile2 = "$env:TEMP\RemoveWorsheet2.xlsx"
            Remove-Item $xlFile2 -ErrorAction SilentlyContinue

            $data | Export-Excel -Path $xlFile2 -WorksheetName Target1
            $data | Export-Excel -Path $xlFile2 -WorksheetName Target2
            $data | Export-Excel -Path $xlFile2 -WorksheetName Target3
            $data | Export-Excel -Path $xlFile2 -WorksheetName Sheet1
        }

        it "Should throw about the Path" {
            {Remove-WorkSheet} | Should throw 'Remove-WorkSheet requires the and Excel file'
        }

        it "Should delete Target2" {
            Remove-WorkSheet -Path $xlFile1 -WorksheetName Target2

            $actual = Get-ExcelSheetInfo -Path $xlFile1

            $actual.Count   | Should Be 3
            $actual[0].Name | Should Be "Target1"
            $actual[1].Name | Should Be "Target3"
            $actual[2].Name | Should Be "Sheet1"
        }

        it "Should delete Sheet1" {
            Remove-WorkSheet -Path $xlFile1

            $actual = Get-ExcelSheetInfo -Path $xlFile1

            $actual.Count   | Should Be 3
            $actual[0].Name | Should Be "Target1"
            $actual[1].Name | Should Be "Target2"
            $actual[2].Name | Should Be "Target3"
        }

        it "Should delete multiple sheets" {
            Remove-WorkSheet -Path $xlFile1 -WorksheetName Target1, Sheet1

            $actual = Get-ExcelSheetInfo -Path $xlFile1

            $actual.Count   | Should Be 2
            $actual[0].Name | Should Be "Target2"
            $actual[1].Name | Should Be "Target3"
        }

        it "Should delete sheet from multiple workbooks" {

            Get-ChildItem "$env:TEMP\RemoveWorsheet*.xlsx" | Remove-WorkSheet

            $actual = Get-ExcelSheetInfo -Path $xlFile1

            $actual.Count   | Should Be 3
            $actual[0].Name | Should Be "Target1"
            $actual[1].Name | Should Be "Target2"
            $actual[2].Name | Should Be "Target3"
        }
    }
}