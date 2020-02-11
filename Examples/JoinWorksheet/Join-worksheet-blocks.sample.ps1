
$path = "$env:TEMP\Test.xlsx"
Remove-item -Path $path -ErrorAction SilentlyContinue
#Export disk volume, and Network adapter to their own sheets.
Get-CimInstance -ClassName Win32_LogicalDisk   |
    Select-Object -Property DeviceId,VolumeName, Size,Freespace |
        Export-Excel -Path $path -WorkSheetname Volumes -NumberFormat "0,000"
Get-NetAdapter  |
    Select-Object -Property Name,InterfaceDescription,MacAddress,LinkSpeed |
        Export-Excel -Path $path -WorkSheetname NetAdapters

#Create a summary page with a title of Summary, label the blocks with the name of the sheet they came from and hide the source sheets
Join-Worksheet -Path $path -HideSource -WorkSheetName Summary -NoHeader -LabelBlocks  -AutoSize -Title "Summary" -TitleBold -TitleSize 22 -show
