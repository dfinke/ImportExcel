$f = ".\testExport.xlsx"

rm $f -ErrorAction Ignore

.\GenDates.ps1 |
    Export-Excel $f -Show -AutoSize -ConditionalText $(
        New-ConditionalText -ConditionalType Last7Days
    )
