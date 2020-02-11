# Get only processes hat have a company name
# Sum up PM by company
# Show the Pie Chart

try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

PieChart -Title "Total PM by Company" `
    (Invoke-Sum (Get-Process|Where-Object company) company pm)

