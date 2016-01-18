$file = ".\conditionalTextFormatting.xlsx"
rm $file -ErrorAction Ignore

Get-Service | 
    Select Status, Name, DisplayName, ServiceName |
    Export-Excel $file -Show -AutoSize -AutoFilter -ConditionalText $(
        New-ConditionalText stop 
        New-ConditionalText runn darkblue cyan
        New-ConditionalText -ConditionalType EndsWith svc wheat green 
        New-ConditionalText -ConditionalType BeginsWith windows darkgreen wheat        
    )