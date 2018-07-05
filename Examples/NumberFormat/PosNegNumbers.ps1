try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$file = "disks.xlsx"

Remove-Item $file -ErrorAction Ignore

$data = $(
    New-PSItem 100 -100
    New-PSItem 1 -1
    New-PSItem 1.2 -1.1
)

$data | Export-Excel -Path $file -Show -AutoSize -NumberFormat "0.#0;-0.#0"