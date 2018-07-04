try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$f = ".\testExport.xlsx"

Remove-Item $f -ErrorAction Ignore

.\GenDates.ps1 |
    Export-Excel $f -Show -AutoSize -ConditionalText $(
        New-ConditionalText -ConditionalType Today
    )
