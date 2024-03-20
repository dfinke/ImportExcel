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

Add-ConditionalFormatting -Worksheet $file.Sheet1 -Range '$1:$1' -ConditionValue '=True' -RuleType Expression -StopIfTrue
Add-ConditionalFormatting -Worksheet $file.Sheet1 -Range '$D:$D' -ConditionValue '=NOT(EXACT($A1,$D1))' -RuleType Expression -ForegroundColor Red
Add-ConditionalFormatting -Worksheet $file.Sheet1 -Range '$E:$E' -ConditionValue '=NOT(EXACT($B1,$E1))' -RuleType Expression -ForegroundColor Red
Add-ConditionalFormatting -Worksheet $file.Sheet1 -Range '$F:$F' -ConditionValue '=NOT(EXACT($C1,$F1))' -RuleType Expression -ForegroundColor Red

Close-ExcelPackage $file -Show
