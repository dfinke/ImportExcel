# Get only processes hat have a company name
# Sum up handles by company
# Show the Pie Chart

try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

PieChart -Title "Total Handles by Company" `
    (Invoke-Sum (Get-Process | Where-Object company) company handles)
