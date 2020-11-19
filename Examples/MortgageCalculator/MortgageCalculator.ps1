<#
  Fixed Rate Loan/Mortgage Calculator in Excel
#>

param(
    $Amount = 400000,
    $InterestRate = .065,
    $Term = 30
)
try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}
function New-CellData {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Does not change system state')]
    param(
        $Range,
        $Value,
        $Format
    )

    $setFormatParams = @{
        Worksheet    = $ws
        Range        = $Range
        NumberFormat = $Format
    }

    if ($Value -is [string] -and $Value.StartsWith('=')) {
        $setFormatParams.Formula = $Value
    }
    else {
        $setFormatParams.Value = $Value
    }

    Set-ExcelRange @setFormatParams
}

$f = "$PSScriptRoot\mortgage.xlsx"
Remove-Item $f -ErrorAction SilentlyContinue

$pkg = "" | Export-Excel $f -Title 'Fixed Rate Loan Payments' -PassThru -AutoSize
$ws = $pkg.Workbook.Worksheets["Sheet1"]

New-CellData -Range A3 -Value 'Amount'
New-CellData -Range B3 -Value $Amount -Format '$#,##0'

New-CellData -Range A4 -Value "Interest Rate"
New-CellData -Range B4 -Value $InterestRate -Format 'Percentage'

New-CellData -Range A5 -Value "Term (Years)"
New-CellData -Range B5 -Value $Term

New-CellData -Range D3 -Value "Monthly Payment"
New-CellData -Range F3 -Value "=-PMT(F4, B5*12, B3)" -Format '$#,##0.#0'

New-CellData -Range D4 -Value "Monthly Rate"
New-CellData -Range F4 -Value "=((1+B4)^(1/12))-1" -Format 'Percentage'

Close-ExcelPackage $pkg -Show