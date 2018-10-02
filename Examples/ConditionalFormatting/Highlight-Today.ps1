try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$f = ".\testExport.xlsx"

Remove-Item -Path $f -ErrorAction Ignore

.\GenDates.ps1 |
    Export-Excel $f -Show -AutoSize -ConditionalText $(
        New-ConditionalText -ConditionalType Today
    )
