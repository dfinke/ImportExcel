$file = "disks.xlsx"

rm $file -ErrorAction Ignore

Get-CimInstance win32_logicaldisk -filter "drivetype=3" | 
    Select DeviceID,Volumename,Size,Freespace | 
    Export-Excel -Path $file -Show -AutoSize -NumberFormat "0"