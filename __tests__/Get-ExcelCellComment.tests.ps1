Describe "Testing comment retrieval " {
    BeforeAll {
		$path = "TestDrive:\test.xlsx"
		
        Remove-Item -path $path -ErrorAction SilentlyContinue
        $excel = ConvertFrom-Csv    @"
Product, City, Gross, Net
Apple, London , 300, 250
Orange, London , 400, 350
Banana, London , 300, 200
Orange, Paris,   600, 500
Banana, Paris,   300, 200
Apple, New York, 1200,700

"@  | Export-Excel  -Path $path  -WorksheetName Sheet1 -PassThru

        Close-ExcelPackage -ExcelPackage $excel
	}
	it "Comment was not present                                                                " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		$comment = Get-ExcelCellComment -Worksheet $ws -Column A -Row 2 
		Close-ExcelPackage -ExcelPackage $excel
		$comment.Text | Should -Be $null
	}
	it "Comment was retrieved                                                                  " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		Set-ExcelCellComment -Worksheet $ws -Column A -Row 2 -Comment "This is a test comment"
		$comment = Get-ExcelCellComment -Worksheet $ws -Column A -Row 2
		Close-ExcelPackage -ExcelPackage $excel
		$comment.Text  | Should      -Be "This is a test comment"
	}
	it "Comment author was retrieved                                                           " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		Set-ExcelCellComment -Worksheet $ws -Column A -Row 3 -Comment "This is a test comment with an author" -Author "Pester"
		$comment = Get-ExcelCellComment -Worksheet $ws -Column A -Row 3
		Close-ExcelPackage -ExcelPackage $excel
		$comment.Author  | Should      -Be "Pester"
	}
}

