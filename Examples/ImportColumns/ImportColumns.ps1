try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

# Create example file
$xlFile = "$PSScriptRoot\ImportColumns.xlsx"
Get-Process | Export-Excel -Path $xlFile
# -ImportColumns will also arrange columns
Import-Excel -Path $xlFile -ImportColumns @(1,3,2) -NoHeader -StartRow 1
# Get only pm, npm, cpu, id, processname
Import-Excel -Path $xlFile -ImportColumns @(6,7,12,25,46) | Format-Table -AutoSize