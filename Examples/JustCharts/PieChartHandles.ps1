# Get only processes hat have a company name
# Sum up handles by company
# Show the Pie Chart

try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

PieChart -Title "Total Handles by Company" `
    (Invoke-Sum (Get-Process | Where-Object company) company handles)
