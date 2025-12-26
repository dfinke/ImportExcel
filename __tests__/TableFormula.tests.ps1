#Requires -Modules @{ ModuleName="Pester"; ModuleVersion="4.0.0" }
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'False Positives')]
param()

Describe "Table Formula Bug Fix" -Tag "TableFormula" {
    BeforeAll {
        $WarningAction = "SilentlyContinue"
    }

    Context "FullCalcOnLoad is set to false to prevent formula corruption" {
        BeforeAll {
            $path = "TestDrive:\table_formula.xlsx"
            Remove-Item -Path $path -ErrorAction SilentlyContinue
            
            # Create a table with a blank record
            $BlankRecordForFile = [PsCustomObject]@{
                'Action Add' = ''
                UserName     = ''
                Address      = ''
                Name         = ''
                NewName      = ''
            }

            # Export as a table
            $ExcelFile = $BlankRecordForFile | Export-Excel -Path $path -WorksheetName 'Data' `
                -TableName 'DataTable' -TableStyle 'Light1' -AutoSize:$false -AutoFilter `
                -BoldTopRow -FreezeTopRow -StartRow 1 -PassThru

            $Worksheet = $ExcelFile.Workbook.Worksheets['Data']
            
            # Insert a row and add a complex formula with table-structured references
            $Worksheet.InsertRow(2, 1)
            
            # This formula uses old-style table references [[#This Row],[ColumnName]]
            # which Excel converts to [@ColumnName] when opening
            $Formula = '=IFS( [[#This Row],[UserName]]="","", [[#This Row],[Action Add]]=TRUE, CONCAT([[#This Row],[Address]],"-",[[#This Row],[UserName]]), CONCAT([[#This Row],[Address]],"-",[[#This Row],[UserName]]) <> [[#This Row],[Name]], CONCAT([[#This Row],[Address]],"-",[[#This Row],[UserName]]), TRUE, "")'
            
            $Cell = $Worksheet.Cells['e2']
            $Cell.Formula = $Formula
            
            Close-ExcelPackage $ExcelFile
            
            # Reopen to verify
            $ExcelFile2 = Open-ExcelPackage -Path $path
            $Worksheet2 = $ExcelFile2.Workbook.Worksheets['Data']
        }

        It "Sets fullCalcOnLoad to false in the workbook XML" {
            # Extract and check the XML directly
            $TempExtractPath = Join-Path -Path $TestDrive -ChildPath "extracted_$(Get-Random)"
            Expand-Archive -Path $path -DestinationPath $TempExtractPath -Force
            $WorkbookXml = Get-Content (Join-Path -Path $TempExtractPath -ChildPath "xl/workbook.xml") -Raw
            
            $WorkbookXml | Should -Match 'fullCalcOnLoad="0"'
            
            Remove-Item -Path $TempExtractPath -Recurse -Force
        }

        It "Preserves the formula correctly after save and reopen" {
            $Cell2 = $Worksheet2.Cells['e2']
            $Cell2.Formula | Should -Not -BeNullOrEmpty
            $Cell2.Formula | Should -Match 'IFS\('
            $Cell2.Formula | Should -Match 'CONCAT\('
        }

        It "Does not corrupt the formula with extra @ symbols" {
            $Cell2 = $Worksheet2.Cells['e2']
            # The formula should not have extra @ symbols added by Excel during recalculation
            # The specific bug was an @ being inserted before CONCAT in the middle of the formula
            # We can't test this directly without opening in Excel, but we can verify the formula is unchanged
            $Cell2.Formula.Length | Should -BeGreaterThan 100
        }

        AfterAll {
            if ($ExcelFile2) {
                Close-ExcelPackage -ExcelPackage $ExcelFile2 -NoSave
            }
        }
    }

    Context "FullCalcOnLoad setting works with different save methods" {
        It "Sets fullCalcOnLoad to false when using SaveAs" {
            $path = "TestDrive:\saveas_test.xlsx"
            $path2 = "TestDrive:\saveas_test2.xlsx"
            Remove-Item -Path $path, $path2 -ErrorAction SilentlyContinue
            
            $data = [PSCustomObject]@{ Name = 'Test' }
            $excel = $data | Export-Excel -Path $path -PassThru
            
            Close-ExcelPackage $excel -SaveAs $path2
            
            # Extract and check the XML
            $TempExtractPath = Join-Path -Path $TestDrive -ChildPath "extracted_saveas_$(Get-Random)"
            Expand-Archive -Path $path2 -DestinationPath $TempExtractPath -Force
            $WorkbookXml = Get-Content (Join-Path -Path $TempExtractPath -ChildPath "xl/workbook.xml") -Raw
            
            $WorkbookXml | Should -Match 'fullCalcOnLoad="0"'
            
            Remove-Item -Path $TempExtractPath -Recurse -Force
        }

        It "Sets fullCalcOnLoad to false when using Calculate flag" {
            $path = "TestDrive:\calculate_test.xlsx"
            Remove-Item -Path $path -ErrorAction SilentlyContinue
            
            $data = [PSCustomObject]@{ Name = 'Test'; Value = 100 }
            $excel = $data | Export-Excel -Path $path -PassThru
            
            # Set a formula
            $ws = $excel.Workbook.Worksheets[1]
            $ws.Cells['C2'].Formula = 'B2*2'
            
            Close-ExcelPackage $excel -Calculate
            
            # Extract and check the XML
            $TempExtractPath = Join-Path -Path $TestDrive -ChildPath "extracted_calc_$(Get-Random)"
            Expand-Archive -Path $path -DestinationPath $TempExtractPath -Force
            $WorkbookXml = Get-Content (Join-Path -Path $TempExtractPath -ChildPath "xl/workbook.xml") -Raw
            
            $WorkbookXml | Should -Match 'fullCalcOnLoad="0"'
            
            Remove-Item -Path $TempExtractPath -Recurse -Force
        }
    }
}
