﻿[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='Set*', Justification='Does not change system state')]
param()

function Set-CellStyle {
    param(
        $WorkSheet,
        $Row,
        $LastColumn,
        [OfficeOpenXml.Style.ExcelFillStyle]$Pattern,
        $Color
    )
    if ($Color -is [string])         {$Color = [System.Drawing.Color]::$Color }
    $t=$WorkSheet.Cells["A$($Row):$($LastColumn)$($Row)"]
    $t.Style.Fill.PatternType=$Pattern
    $t.Style.Fill.BackgroundColor.SetColor($Color)
}