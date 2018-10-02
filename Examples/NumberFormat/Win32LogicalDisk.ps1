try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$file = "disks.xlsx"

Remove-Item -Path $file -ErrorAction Ignore

Get-CimInstance -ClassName win32_logicaldisk -filter "drivetype=3" |
    Select-Object -Property DeviceID,Volumename,Size,Freespace |
    Export-Excel -Path $file -Show -AutoSize -NumberFormat "0"