param ($fibonacciDigits=10)

try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

$(
    New-PSItem 0
    New-PSItem 1

    (
        2..$fibonacciDigits |
            ForEach-Object {
                New-PSItem ('=a{0}+a{1}' -f ($_+1),$_)
            }
    )
) | Export-Excel $xlSourcefile -Show
