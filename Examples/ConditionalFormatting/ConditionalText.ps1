$f = ".\conditionalTextFormatting.xlsx"
rm $f -ErrorAction Ignore

Get-Service | 
    Select Status, Name, DisplayName, ServiceName |
    Export-Excel $f -Show -AutoSize -ConditionalText $(
        New-ConditionalText stop darkred
        New-ConditionalText running darkblue
        New-ConditionalText app DarkMagenta
    )