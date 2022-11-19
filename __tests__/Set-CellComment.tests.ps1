if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}

Describe "Test setting comment on cells in different ways" -Tag SetCellComment {
    BeforeAll {
        $data = ConvertFrom-Csv @"
OrderId,Category,Sales,Quantity,Discount
1,Cosmetics,744.01,07,0.7
2,Grocery,349.13,25,0.3
3,Apparels,535.11,88,0.2
4,Electronics,524.69,60,0.1
5,Electronics,439.10,41,0.0
6,Apparels,56.84,54,0.8
7,Electronics,326.66,97,0.7
8,Cosmetics,17.25,74,0.6
9,Grocery,199.96,39,0.4
10,Grocery,731.77,20,0.3
"@

        $Excel = $data | Export-Excel -PassThru
        $ws = $Excel.Workbook.Worksheets | Select-Object -First 1
    }

    AfterAll {
        Close-ExcelPackage $Excel
    }

    It "Should add comments to multiple cells".PadRight(87) {
        Set-CellComment -Range "A1" -Worksheet $ws -Text "This was added with a single cell range"
        Set-CellComment -Range "A2:C2" -Worksheet $ws -Text "This was added with a multiple cell range"
        Set-CellComment -ColumnLetter A -Row 3 -Worksheet $ws -Text "This was added using a column letter and rownumber"
        Set-CellComment -ColumnNumber 1 -Row 4 -Worksheet $ws -Text "This was added using a column number and row number"

        Set-CellComment -Range "B2" -Worksheet $ws -Text "This demonstrates an overwrite of a previously set comment"

        $ws.Cells["A1"].Comment.Text | Should -BeExactly "This was added with a single cell range"
        $ws.Cells["A2"].Comment.Text | Should -BeExactly "This was added with a multiple cell range" 
        $ws.Cells["B2"].Comment.Text | Should -BeExactly "This demonstrates an overwrite of a previously set comment" 
        $ws.Cells["C2"].Comment.Text | Should -BeExactly "This was added with a multiple cell range"
        $ws.Cells["A3"].Comment.Text | Should -BeExactly "This was added using a column letter and rownumber"
        $ws.Cells["A4"].Comment.Text | Should -BeExactly "This was added using a column number and row number"
    }
}