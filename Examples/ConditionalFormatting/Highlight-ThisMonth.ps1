try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$f = ".\testExport.xlsx"

Remove-Item $f -ErrorAction Ignore

.\GenDates.ps1 |
    Export-Excel $f -Show -AutoSize -ConditionalText $(
        New-ConditionalText -ConditionalType ThisMonth
    )
