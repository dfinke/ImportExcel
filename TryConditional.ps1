$xlFile = "c:\temp\testCF.xlsx"
rm $xlFile -ErrorAction Ignore

$data = Get-Process | where Company | select Company,pm,handles,name

$cfHandles = New-ConditionalFormattingIconSet `
    -Address "C:C" `
    -ConditionalFormat ThreeIconSet `
    -IconType Flags -Reverse

$cfPM = New-ConditionalFormattingIconSet `
    -Address "B:B" `
    -ConditionalFormat FourIconSet `
    -Reverse -IconType TrafficLights


$data |
    Export-Excel -AutoSize -AutoFilter -ConditionalFormat $cfHandles, $cfPM -Path $xlFile -Show