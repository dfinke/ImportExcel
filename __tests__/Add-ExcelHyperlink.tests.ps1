Describe "Testing adding hyperlink" {
    BeforeAll {
		$path = "TestDrive:\addhyperlink.xlsx"
        $path = "$ENV:TEMP\addhyperlink.xlsx"
        $worksheetName = "Sheet1"
        $cell = "A2"
        $cell2 = "A3"
        $hyperlink = "NamedRange"

        Remove-Item -Path $path -ErrorAction SilentlyContinue
        1..5 | Export-Excel -Path $path -WorksheetName $worksheetName 

        $excelPackage = Open-ExcelPackage -Path $path
        $excel = $excelPackage.Workbook.Worksheets[$worksheetName]

        $rangeCells=$excel.Cells["D1:E12"]
        $excelPackage.Workbook.Names.Add($rangeName,$rangeCells)

        Close-ExcelPackage -ExcelPackage $excelPackage
	}
    AfterAll {
        $excelPackage = Open-ExcelPackage -Path $path -KillExcel
        Close-ExcelPackage -ExcelPackage $excelPackage -NoSave
        Remove-Item -Path $path -ErrorAction SilentlyContinue
    
    }
	It "Hyperlink does not exist" {
		$hyperlink = Get-ExcelHyperlink -Path $path
		$hyperlink | Should -Be $null

	}
	It "Hyperlink is added" {
		$null = Add-ExcelHyperlink -Path $path -WorksheetName $worksheetName -Cell $cell -Hyperlink $rangeName -DisplayName 'display text'

		$hyperlink2 = Get-ExcelHyperlink -Path $path -WorksheetName $worksheetName -Cell $cell
        ($hyperlink2).Address| Should -Be $Cell
	}

	It "First cell's style with hyperlink is changed from Normal to Hyperlink" {
        ($hyperlink2).StyleName| Should -Be 'Hyperlink'
	}

	It "Next cell's style with hyperlink is changed from Normal to Hyperlink" {
        $null = Add-ExcelHyperlink -Path $path -WorksheetName $worksheetName -Cell $cell2 -Hyperlink $rangeName -DisplayName 'display text'

        $hyperlink3 = Get-ExcelHyperlink -Path $path -WorksheetName $worksheetName -Cell $cell2
        ($hyperlink3).StyleName| Should -Be 'Hyperlink'
	}
}