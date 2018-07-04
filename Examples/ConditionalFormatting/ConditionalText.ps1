try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$file = ".\conditionalTextFormatting.xlsx"
Remove-Item $file -ErrorAction Ignore

Get-Service |
    Select-Object Status, Name, DisplayName, ServiceName |
    Export-Excel $file -Show -AutoSize -AutoFilter -ConditionalText $(
        New-ConditionalText stop
        New-ConditionalText runn darkblue cyan
        New-ConditionalText -ConditionalType EndsWith svc wheat green
        New-ConditionalText -ConditionalType BeginsWith windows darkgreen wheat
    )