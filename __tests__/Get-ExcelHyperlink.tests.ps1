Describe "Testing getting hyperlinks" {
        BeforeAll {
                    $path = "TestDrive:\gethyperlink.xlsx"
            $path = "$ENV:TEMP\gethyperlink.xlsx"
            $worksheetName = "Sheet1"
            $worksheetName2 = "Sheet2"
    
            $cell = "A2"
            $cell2 = "B3"
            $hyperlink = "NamedRange"
    
            Remove-Item -Path $path -ErrorAction SilentlyContinue
            1..5 | Export-Excel -Path $path -WorksheetName $worksheetName 
    
            $excelPackage = Open-ExcelPackage -Path $path
            Copy-ExcelWorksheet -SourceWorkbook $excelPackage -DestinationWorkbook $excelPackage -DestinationWorksheet $worksheetName2
    
            $excel = $excelPackage.Workbook.Worksheets[$worksheetName]
    
            $rangeCells=$excel.Cells["D1:E12"]
            $excelPackage.Workbook.Names.Add($rangeName,$rangeCells)
    
            Close-ExcelPackage -ExcelPackage $excelPackage
    
            }
        
        It "Hyperlink does not exist in a sheet" {
                    $hyperlink = Get-ExcelHyperlink -Path $path -WorksheetName $worksheetName
                    $hyperlink.Worksheet | Should -Be $null
            }
    
        It "Hyperlink exists in a specified sheet" {
            $null = Add-ExcelHyperlink -Path $path -WorksheetName $worksheetName2 -Cell $cell2 -Hyperlink $rangeName -DisplayName 'display text'
                    $hyperlink2 = Get-ExcelHyperlink -Path $path -WorksheetName $worksheetName2
                    $hyperlink2.Worksheet | Should -Be $worksheetName2
            }
    
            It "Hyperlink exists in a specified cell" {
                    $hyperlink3 = Get-ExcelHyperlink -Path $path -WorksheetName $worksheetName2 -Cell $cell2
                    $hyperlink3.Address | Should -Be $cell2
            }
    
        It "Hyperlink exists in multiple sheets" {
            $null = Add-ExcelHyperlink -Path $path -WorksheetName $worksheetName -Cell $cell -Hyperlink $rangeName -DisplayName 'display text 2'
                    $hyperlink4 = Get-ExcelHyperlink -Path $path
                    $hyperlink4.Count | Should -Be 2
            }
    
    }