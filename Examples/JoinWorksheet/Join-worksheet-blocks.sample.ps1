try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

#Export disk volume, and Network adapter to their own sheets.
Get-CimInstance -ClassName Win32_LogicalDisk   |
    Select-Object -Property DeviceId,VolumeName, Size,Freespace |
        Export-Excel -Path $xlSourcefile -WorkSheetname Volumes -NumberFormat "0,000"
Get-NetAdapter  |
    Select-Object -Property Name,InterfaceDescription,MacAddress,LinkSpeed |
        Export-Excel -Path $xlSourcefile -WorkSheetname NetAdapters

#Create a summary page with a title of Summary, label the blocks with the name of the sheet they came from and hide the source sheets
Join-Worksheet -Path $xlSourcefile -HideSource -WorkSheetName Summary -NoHeader -LabelBlocks  -AutoSize -Title "Summary" -TitleBold -TitleSize 22 -show
