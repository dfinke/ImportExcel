try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$file = "disks.xlsx"

Remove-Item $file -ErrorAction Ignore

Get-CimInstance win32_logicaldisk -filter "drivetype=3" |
    Select-Object DeviceID,Volumename,Size,Freespace |
    Export-Excel -Path $file -Show -AutoSize -NumberFormat "0"