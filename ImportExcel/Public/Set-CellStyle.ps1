function Set-CellStyle {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Does not change system state')]
    [CmdletBinding()]
    param(
        $Worksheet,
        $Row,
        $LastColumn,
        [OfficeOpenXml.Style.ExcelFillStyle]$Pattern,
        $Color
    )
    if ($Color -is [string])         {$Color = [System.Drawing.Color]::$Color }
    $t=$Worksheet.Cells["A$($Row):$($LastColumn)$($Row)"]
    $t.Style.Fill.PatternType=$Pattern
    $t.Style.Fill.BackgroundColor.SetColor($Color)
}