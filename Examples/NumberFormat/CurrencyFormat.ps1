try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$file = "$env:temp\disks.xlsx"

Remove-Item -Path $file -ErrorAction Ignore

$data = $(
    New-PSItem 100 -100
    New-PSItem 1 -1
    New-PSItem 1.2 -1.1
    New-PSItem -3.2 -4.1
    New-PSItem -5.2 6.1
    New-PSItem 1000 -2000
)
#Number format can expand terms like Currency, to the local currency format
$data | Export-Excel -Path $file -Show -AutoSize -NumberFormat 'Currency'
