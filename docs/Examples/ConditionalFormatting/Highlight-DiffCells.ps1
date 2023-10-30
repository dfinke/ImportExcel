try { Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 } catch { throw ; return }

$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"

Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

$data = ConvertFrom-Csv @"
Region,State,Units2021,Units2022
West,Texas,927,925
North,Tennessee,466,466
East,Florida,520,458
East,Maine,828,661
West,Virginia,465,465
North,Missouri,436,235
South,Kansas,214,214
North,North Dakota,789,640
South,Delaware,712,508
"@

$excel = $data | Export-Excel $xlSourcefile -AutoSize -PassThru

Add-ConditionalFormatting -Worksheet $excel.sheet1 -Range "C2:D10" -ConditionValue '=$C2=$D2' -RuleType Expression -BackgroundColor ([System.Drawing.Color]::Thistle) -Bold
Add-ConditionalFormatting -Worksheet $excel.sheet1 -Range "A2:D10" -ConditionValue '=$C2=$D2' -RuleType Expression -BackgroundColor ([System.Drawing.Color]::LavenderBlush)

Close-ExcelPackage $excel -Show