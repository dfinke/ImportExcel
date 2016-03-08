param ($fibonacciDigits=10)

$file = "fib.xlsx"
rm "fib.xlsx" -ErrorAction Ignore

$(
    New-PSItem 0
    New-PSItem 1
    
    (
        2..$fibonacciDigits |
            ForEach {
                New-PSItem ('=a{0}+a{1}' -f ($_+1),$_)
            }
    )
) | Export-Excel $file -Show
