#Requires -Modules Pester
Import-Module $PSScriptRoot\..\ImportExcel.psd1 -Force

Describe "Remove Worksheet" {
    Context "Remove a worksheet output" {
        BeforeAll {
            # Create three sheets
            $data = ConvertFrom-Csv @"
"@
            $xlFile = "$env:TEMP\RemoveWorsheet.xlsx"
            Remove-Item $xlFile -ErrorAction SilentlyContinue

            $data | Export-Excel -Path $xlFile -WorksheetName Target1
            $data | Export-Excel -Path $xlFile -WorksheetName Target2
            $data | Export-Excel -Path $xlFile -WorksheetName Target3
        }

        it "Should delete Target2" {
            Remove-WorkSheet -Path $xlFile -WorksheetName Target2

            $actual = Get-ExcelSheetInfo -Path $xlFile

            $actual.Count   | Should Be 2
            $actual[0].Name | Should Be "Target1"
            $actual[1].Name | Should Be "Target3"
        }
    }
}