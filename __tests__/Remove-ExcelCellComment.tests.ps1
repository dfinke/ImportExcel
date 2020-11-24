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
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
	}
	it "Comment was not present                                                                " {
		Get-ExcelCellComment -Worksheet $ws -Column A -Row 2 | Should -Be $null
	}
	it "Comment is present                                                                     " {
		Set-ExcelCellComment -Worksheet $ws -Column A -Row 2 -Comment "This is a test comment"
		$comment = Get-ExcelCellComment -Worksheet $ws -Column A -Row 2
		$comment.Text  | Should      -Be "This is a test comment"
	}
	it "Comment was removed                                                                    " {
		Remove-ExcelCellComment -Worksheet $ws -Column A -Row 2
		Get-ExcelCellComment -Worksheet $ws -Column A -Row 2  | Should      -Be $null
	}
}

