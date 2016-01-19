echo Last7Days LastMonth LastWeek NextMonth NextWeek ThisMonth ThisWeek Today Tomorrow Yesterday |
    % {
    $text = @"
`$f = ".\testExport.xlsx"

rm `$f -ErrorAction Ignore

.\GenDates.ps1 |
    Export-Excel `$f -Show -AutoSize -ConditionalText `$(
        New-ConditionalText -ConditionalType $_
    )
"@
        $text | Set-Content -Encoding Ascii "Highlight-$($_).ps1"
    }
