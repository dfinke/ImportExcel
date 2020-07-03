try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

#Put some simple data in a worksheet and Get an excel package object to represent the file
1..5 | Export-Excel $xlSourcefile -WorksheetName 'Tab1' -AutoSize -AutoFilter

#Add another tab.  Replace the $TabData2 with your data
1..10 | Export-Excel $xlSourcefile -WorksheetName 'Tab 2' -AutoSize -AutoFilter

#Add another tab.  Replace the $TabData3 with your data
1..15  | Export-Excel $xlSourcefile -WorksheetName 'Tab 3' -AutoSize -AutoFilter -Show
