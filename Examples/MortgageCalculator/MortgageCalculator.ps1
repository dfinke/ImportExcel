<#
  Fixed Rate Loan/Mortgage Calculator in Excel
#>

$f = "$PSScriptRoot\mortgage.xlsx"
rm $f -ErrorAction SilentlyContinue

$pkg = "" | Export-Excel $f -Title 'Fixed Rate Loan Payments' -PassThru -AutoSize

$ws = $pkg.Workbook.Worksheets["Sheet1"]

Set-Format -WorkSheet $ws -Range "A3" -Value "Amount"
Set-Format -WorkSheet $ws -Range "B3" -Value 400000 -NumberFormat '$#,##0'

Set-Format -WorkSheet $ws -Range "A4" -Value "Interest Rate"
Set-Format -WorkSheet $ws -Range "B4" -Value .065 -NumberFormat 'Percentage'

Set-Format -WorkSheet $ws -Range "A5" -Value "Term (Years)"
Set-Format -WorkSheet $ws -Range "B5" -Value 30 

Set-Format -WorkSheet $ws -Range "D3" -Value "Monthly Payment" 
Set-Format -WorkSheet $ws -Range "F3" -Formula "=-PMT(F4, B5*12, B3)" -NumberFormat '$#,##0.#0'

Set-Format -WorkSheet $ws -Range "D4" -Value "Monthly Rate"
Set-Format -WorkSheet $ws -Range "F4" -Formula "=((1+B4)^(1/12))-1" -NumberFormat 'Percentage'

Close-ExcelPackage $pkg -Show

