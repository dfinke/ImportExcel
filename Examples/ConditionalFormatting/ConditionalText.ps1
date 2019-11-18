try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$file = "$env:temp\conditionalTextFormatting.xlsx"
Remove-Item $file -ErrorAction Ignore

Get-Service |
    Select-Object Status, Name, DisplayName, ServiceName |
    Export-Excel $file -Show -AutoSize -AutoFilter -ConditionalText $(
        New-ConditionalText stop                                                  #Stop is the condition value, the rule is defaults to 'Contains text' and the default Colors are used
        New-ConditionalText runn darkblue cyan                                    #runn is the condition value, the rule is defaults to 'Contains text'; the foregroundColur is darkblue and the background is cyan
        New-ConditionalText -ConditionalType EndsWith svc wheat green             #the rule here is 'Ends with' and the value is 'svc' the forground is wheat and the background dark green
        New-ConditionalText -ConditionalType BeginsWith windows darkgreen wheat   #this is 'Begins with "Windows"' the forground is dark green and the background wheat
    )