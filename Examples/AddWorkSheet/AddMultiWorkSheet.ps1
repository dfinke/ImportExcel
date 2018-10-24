try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$xlSourcefile = "$env:TEMP\Source.xlsx"
write-host "Save location: $xlSourcefile"

Remove-Item $xlSourcefile -ErrorAction Ignore

#Put some simple data in a worksheet and Get an excel package object to represent the file
$TabData1 = 1..5 | Export-Excel $xlSourcefile -WorksheetName 'Tab 1' -AutoSize -AutoFilter

#Add another tab.  Replace the $TabData2 with your data
$TabData2 = 1..10 | Export-Excel $xlSourcefile -WorksheetName 'Tab 2' -AutoSize -AutoFilter

#Add another tab.  Replace the $TabData3 with your data
$TabData3 = 1..15  | Export-Excel $xlSourcefile -WorksheetName 'Tab 3' -AutoSize -AutoFilter -Show
