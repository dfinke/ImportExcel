try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$file = "$env:TEMP\disks.xlsx"

Remove-Item $file -ErrorAction Ignore

$data = $(
    New-PSItem 100 -100
    New-PSItem 1 -1
    New-PSItem 1.2 -1.1
    New-PSItem -3.2 -4.1
    New-PSItem -5.2 6.1
)
#Set the numbers throughout the sheet to format as positive in blue with a + sign, negative in Red with a - sign.
$data | Export-Excel -Path $file -Show -AutoSize -NumberFormat "[Blue]+0.#0;[Red]-0.#0"
