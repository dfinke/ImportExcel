<#
  Fixed Rate Loan/Mortgage Calculator in Excel
#>

param(
    $Amount = 400000,
    $InterestRate = .065,
    $Term = 30
)

function New-CellData {
    param(
        $Range,
        $Value,
        $Format
    )

    $setFormatParams = @{
        WorkSheet    = $ws
        Range        = $Range
        NumberFormat = $Format
    }

    if ($Value -is [string] -and $Value.StartsWith('=')) {
        $setFormatParams.Formula = $Value
    }
    else {
        $setFormatParams.Value = $Value
    }

    Set-Format @setFormatParams
}

$f = "$PSScriptRoot\mortgage.xlsx"
Remove-Item $f -ErrorAction SilentlyContinue

$pkg = "" | Export-Excel $f -Title 'Fixed Rate Loan Payments' -PassThru -AutoSize
$ws = $pkg.Workbook.Worksheets["Sheet1"]

New-CellData A3 'Amount'
New-CellData B3 $Amount '$#,##0'

New-CellData A4 "Interest Rate"
New-CellData B4 $InterestRate 'Percentage'

New-CellData A5 "Term (Years)"
New-CellData B5 $Term

New-CellData D3 "Monthly Payment"
New-CellData F3 "=-PMT(F4, B5*12, B3)" '$#,##0.#0'

New-CellData D4 "Monthly Rate"
New-CellData F4 "=((1+B4)^(1/12))-1" 'Percentage'

Close-ExcelPackage $pkg -Show