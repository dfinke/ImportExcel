"Last7Days", "LastMonth", "LastWeek", "NextMonth", "NextWeek", "ThisMonth", "ThisWeek", "Today", "Tomorrow", "Yesterday" |
    Foreach-Object {
    $text = @"
`$f = ".\testExport.xlsx"

remove-item `$f -ErrorAction Ignore

.\GenDates.ps1 |
    Export-Excel `$f -Show -AutoSize -ConditionalText `$(
        New-ConditionalText -ConditionalType $_
    )
"@
        $text | Set-Content -Encoding Ascii "Highlight-$($_).ps1"
    }
