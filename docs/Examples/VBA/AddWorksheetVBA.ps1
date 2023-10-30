$xlfile = "$env:temp\test.xlsm"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$Excel = ConvertFrom-Csv @"
Region,Item,TotalSold
West,screwdriver,98
West,kiwi,19
North,kiwi,47
West,screws,48
West,avocado,52
East,avocado,40
South,drill,61
North,orange,92
South,drill,29
South,saw,36
"@ | Export-Excel $xlfile -TableName 'Sales' -WorksheetName 'Sales' -AutoSize -PassThru

$wb = $Excel.Workbook
$sheet = $wb.Worksheets["Sales"]
$wb.CreateVBAProject()

# Add a sub to the 'Worksheet_SelectionChange' event of the worksheet to highlight the selected row & column of the active table.
# https://docs.microsoft.com/en-gb/office/vba/excel/Concepts/Cells-and-Ranges/highlight-the-active-cell-row-or-column
$code = @"
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Clear the color of all the cells
    Cells.Interior.ColorIndex = 0
    If Target.Cells.Count > 1 Then Exit Sub
        Application.ScreenUpdating = False
        With ActiveCell
            ' Highlight the row and column that contain the active cell, within the current region
            Range(Cells(.Row, .CurrentRegion.Column), Cells(.Row, .CurrentRegion.Columns.Count + .CurrentRegion.Column - 1)).Interior.ColorIndex = 38
            Range(Cells(.CurrentRegion.Row, .Column), Cells(.CurrentRegion.Rows.Count + .CurrentRegion.Row - 1, .Column)).Interior.ColorIndex = 24
        End With
    Application.ScreenUpdating = True
End Sub
"@

$sheet.CodeModule.Code = $code

Close-ExcelPackage $Excel -Show