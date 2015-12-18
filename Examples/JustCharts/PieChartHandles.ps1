# Get only processes hat have a company name
# Sum up handles by company
# Show the Pie Chart

PieChart -Title "Total Handles by Company" `
    (Invoke-Sum (Get-Process|Where company) company handles)
