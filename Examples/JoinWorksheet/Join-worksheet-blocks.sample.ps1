
$path = "$env:TEMP\Test.xlsx"
Remove-item -Path $path -ErrorAction SilentlyContinue
 Get-WmiObject -Class win32_logicaldisk |
    Select-Object -Property DeviceId,VolumeName, Size,Freespace | 
        Export-Excel -Path $path -WorkSheetname Volumes -NumberFormat "0,000"
Get-NetAdapter  | 
    Select-Object -Property Name,InterfaceDescription,MacAddress,LinkSpeed | 
        Export-Excel -Path $path -WorkSheetname NetAdapters

Join-Worksheet -Path $path -HideSource -WorkSheetName Summary -NoHeader -LabelBlocks  -AutoSize -Title "Summary" -TitleBold -TitleSize 22 -show  
