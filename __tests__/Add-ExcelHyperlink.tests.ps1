Describe "Testing adding hyperlink" {
    BeforeAll {
		$path = "TestDrive:\hyperlink.xlsx"
        $worksheetName = "Sheet1"
        $cell = "A2"
        $hyperlink = "NamedRange"

        Remove-Item -Path $path -ErrorAction SilentlyContinue
        1..5 | Export-Excel -Path $path -WorksheetName $worksheetName 

        $excelPackage = Open-ExcelPackage -Path $path
        $excel = $excelPackage.Workbook.Worksheets[$worksheetName]

        $rangeCells=$excel.Cells["D1:E12"]
        $excelPackage.Workbook.Names.Add($rangeName,$rangeCells)

        Close-ExcelPackage -ExcelPackage $excelPackage
	}
	It "Hyperlink does not exist" {
		$hyperlink = Get-ExcelHyperlink -Path $path -WorksheetName $worksheetName  -Cell $Cell
		$hyperlink | Should -Be $null

	}
	It "Hyperlink is been added" {
		Add-ExcelHyperlink -Path $path -WorksheetName $worksheetName -Cell $Cell -Hyperlink $rangeName -DisplayName 'display text'

		$hyperlink2 = Get-ExcelHyperlink -Path $path -WorksheetName $worksheetName -Cell $Cell 
        ($hyperlink2).Address| Should -Be $Cell
	}

	It "Cell style is changed from Normal to Hyperlink" {
        ($hyperlink2).StyleName| Should -Be 'Hyperlink'
	}
}