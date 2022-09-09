try { Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 } catch { throw ; return }

## This exports only the numbers
# 1..10 | Export-excel $PSScriptRoot\test.xlsx -Show

## This exports the numbers and in A1 the text "MyNum"
# 1..10 | Export-excel $PSScriptRoot\test.xlsx -HeaderName MyNum -Show

$xlfile = "$PSScriptRoot/testMultipleColumns.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$Regions = 'West', 'North', 'East ', 'East ', 'West ', 'North', 'South', 'North', 'South'
$States = 'Texas', 'Tennessee', 'Florida', 'Maine', 'Virginia', 'Missouri', 'Kansas', 'North Dakota', 'Delaware'
$Units = 927, 466, 520, 828, 465, 436, 214, 789, 712
$Prices = 923.71, 770.67, 458.68, 661.24, 53.58, 235.67, 992.47, 640.72, 508.55

# Export each list (array) as a separate column to the same worksheet and workbook
$Regions | Export-Excel -Path $xlfile -HeaderName Region -StartColumn 1 -AutoSize
$States  | Export-Excel -Path $xlfile -HeaderName State  -StartColumn 2 -AutoSize
$Units   | Export-Excel -Path $xlfile -HeaderName Units  -StartColumn 3 -AutoSize
$Prices  | Export-Excel -Path $xlfile -HeaderName Prices -StartColumn 4 -AutoSize

# Show the results in Excel
. $xlfile