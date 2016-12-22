function Set-CellStyle {
    param(
        $WorkSheet,
        $Row,
        $LastColumn,
        [OfficeOpenXml.Style.ExcelFillStyle]$Pattern,
        [System.Drawing.Color]$Color
    )

    $t=$WorkSheet.Cells["A$($Row):$($LastColumn)$($Row)"]
    $t.Style.Fill.PatternType=$Pattern
    $t.Style.Fill.BackgroundColor.SetColor($Color)
}