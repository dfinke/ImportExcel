try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$file = "disks.xlsx"

Remove-Item $file -ErrorAction Ignore

Get-CimInstance win32_logicaldisk -filter "drivetype=3" |
    Select-Object DeviceID,Volumename,Size,Freespace |
    Export-Excel -Path $file -Show -AutoSize