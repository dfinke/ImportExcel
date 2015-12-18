# Get only processes hat have a company name
# Sum up PM by company
# Show the Pie Chart

PieChart -Title "Total PM by Company" `
    (Invoke-Sum (Get-Process|Where company) company pm)

