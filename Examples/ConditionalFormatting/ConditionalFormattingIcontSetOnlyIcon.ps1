try { Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 } catch { throw ; return }

$data = ConvertFrom-Csv @"
Region,State,Other,Units,Price,InStock
West,Texas,1,927,923.71,1
North,Tennessee,3,466,770.67,0
East,Florida,0,1520,458.68,1
East,Maine,1,1828,661.24,0
West,Virginia,1,465,053.58,1
North,Missouri,1,436,235.67,1
South,Kansas,0,214,992.47,1
North,North Dakota,1,789,640.72,0 
South,Delaware,-1,712,508.55,1
"@

$xlfile = "$PSScriptRoot\test.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$cfi1 = New-ConditionalFormattingIconSet -Range C:C -ConditionalFormat ThreeIconSet -IconType Symbols -ShowIconOnly
$cfi2 = New-ConditionalFormattingIconSet -Range F:F -ConditionalFormat ThreeIconSet -IconType Symbols2 -ShowIconOnly

$data | Export-Excel $xlfile -AutoSize -ConditionalFormat $cfi1, $cfi2 -Show