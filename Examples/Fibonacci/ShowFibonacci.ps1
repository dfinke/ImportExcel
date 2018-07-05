param ($fibonacciDigits=10)

try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$file = "fib.xlsx"
Remove-Item "fib.xlsx" -ErrorAction Ignore

$(
    New-PSItem 0
    New-PSItem 1

    (
        2..$fibonacciDigits |
            ForEach-Object {
                New-PSItem ('=a{0}+a{1}' -f ($_+1),$_)
            }
    )
) | Export-Excel $file -Show
