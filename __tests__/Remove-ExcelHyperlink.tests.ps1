Describe "Testing removing hyperlinks" {
    BeforeAll {
        $path = "TestDrive:\removehyperlink.xlsx"
        $worksheetName = "Sheet1"
        $worksheetName2 = "Sheet2"

        $cell = "A2"
        $cell2 = "B3"
        $rangeName = "NamedRange"

        

        Remove-Item -Path $path -ErrorAction SilentlyContinue
        1..5 | Export-Excel -Path $path -WorksheetName $worksheetName 

        $excelPackage = Open-ExcelPackage -Path $path
        Copy-ExcelWorksheet -SourceWorkbook $excelPackage -DestinationWorkbook $excelPackage -DestinationWorksheet $worksheetName2

        $excel = $excelPackage.Workbook.Worksheets[$worksheetName]
        

        $rangeCells=$excel.Cells["D1:E12"]
        $excelPackage.Workbook.Names.Add($rangeName,$rangeCells)

        Close-ExcelPackage -ExcelPackage $excelPackage

        $null = Add-ExcelHyperlink -Path $path -WorksheetName $worksheetName -Cell $cell -Hyperlink $rangeName -DisplayName 'Link to something'
        $null = Add-ExcelHyperlink -Path $path -WorksheetName $worksheetName -Cell $cell2 -Hyperlink $rangeName -DisplayName 'Link to something'
        $null = Add-ExcelHyperlink -Path $path -WorksheetName $worksheetName2 -Cell $cell2 -Hyperlink $rangeName
        $null = Add-ExcelHyperlink -Path $path -WorksheetName $worksheetName2 -Cell $cell -Hyperlink $rangeName
        }

    AfterAll {
        $excelPackage = Open-ExcelPackage -Path $path -KillExcel
        Close-ExcelPackage -ExcelPackage $excelPackage -NoSave
        Remove-Item -Path $path -ErrorAction SilentlyContinue
    }
    
    It "Hyperlinks exist in a workbook" {
            $hyperlink = Get-ExcelHyperlink -Path $path
            $hyperlink.Count | Should -Be 4
        }

    It "Hyperlinks exist in a specified sheet" {
            $hyperlink2 = Get-ExcelHyperlink -Path $path -WorksheetName $worksheetName2
            $hyperlink2.Count | Should -Be 2
        }
       
    It "Hyperlink exists in a specified cell" {
            $hyperlink3 = Get-ExcelHyperlink -Path $path -WorksheetName $worksheetName -Cell $cell
            $hyperlink3.Hyperlink | Should -Not -BeNullOrEmpty
        }

    It "Hyperlink from the cell is removed" {
            $null = Remove-ExcelHyperlink -Path $path -WorksheetName $worksheetName -Cell $cell
            $hyperlink3 = Get-ExcelHyperlink -Path $path -WorksheetName $worksheetName -Cell $cell
            $hyperlink3.Hyperlink | Should -Be $null
        }

    It "StyleName from the cell is reset to Normal" {
            $hyperlink3 = Get-ExcelHyperlink -Path $path -WorksheetName $worksheetName -Cell $cell
            $hyperlink3.StyleName | Should -Be 'Normal'
        }

    It "Hyperlinks from the worksheet are removed" {
            $null = Remove-ExcelHyperlink -Path $path -WorksheetName $worksheetName
            $hyperlink4 = Get-ExcelHyperlink -Path $path -WorksheetName $worksheetName
            $hyperlink4.Count | Should -Be 0
     }

    It "Hyperlinks from the workbook are removed" {
            $null = Remove-ExcelHyperlink -Path $path
            $hyperlink5 = Get-ExcelHyperlink -Path $path
            $hyperlink5.Count | Should -Be 0
     }
}