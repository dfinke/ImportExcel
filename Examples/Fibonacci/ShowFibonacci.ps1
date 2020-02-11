param ($fibonacciDigits=10)

try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

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
