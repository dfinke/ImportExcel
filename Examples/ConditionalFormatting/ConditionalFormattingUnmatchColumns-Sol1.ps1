try { Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 } catch { throw ; return }

$data = ConvertFrom-Csv @"
Sequence1,StateName1,Abbreviation1,Sequence2,StateName2,Abbreviation2
A001,Alabama,AL,A001,Alabama,AL
A002,California,CA,B002,California,CA
A003,Colorado,CO,A003,Colorado,CO
A004,Florida,FL,A004,Maine,FL
A005,Missouri,MO,A005,Missouri,MI
A006,New Mexico,NM,A006,Tennessee,TN
A007,North Carolina,NC,C007,North Carolina,NC
A008,North Dakota,ND,A008,South Dakota,SD
A009,Rhode Island,RI,D009,Rhode Island,VA
A010,Texas,TX,A010,Texas,TX
"@

$xlfile = "$PSScriptRoot\test.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$file = $data | Export-Excel $xlfile -AutoSize -PassThru

Add-ConditionalFormatting -Worksheet $file.sheet1 -Range '$D$2:$D$1048576' -ConditionValue '=NOT(EXACT($A2,$D2))' -RuleType Expression -ForegroundColor Red
Add-ConditionalFormatting -Worksheet $file.sheet1 -Range '$E$2:$E$1048576' -ConditionValue '=NOT(EXACT($B2,$E2))' -RuleType Expression -ForegroundColor Red
Add-ConditionalFormatting -Worksheet $file.sheet1 -Range '$F$2:$F$1048576' -ConditionValue '=NOT(EXACT($C2,$F2))' -RuleType Expression -ForegroundColor Red

Close-ExcelPackage $file -Show
