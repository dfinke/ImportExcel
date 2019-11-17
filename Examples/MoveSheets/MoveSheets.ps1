try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$xlfile = "$env:TEMP\testThis.xlsx"
Remove-Item $xlfile -ErrorAction Ignore

1..10  | Export-Excel $xlfile -WorkSheetname First                           #'First' will be the only sheet
11..20 | Export-Excel $xlfile -WorkSheetname Second -MoveToStart             #'Second' is moved before first so the order is 'Second', 'First'
21..30 | Export-Excel $xlfile -WorkSheetname Third  -MoveBefore First        #'Second' is moved before first so the order is 'Second', 'Third', 'First'
31..40 | Export-Excel $xlfile -WorkSheetname Fourth -MoveAfter Third -Show   #'Fourth' is moved after third so the order is ' 'Second', 'Third', 'Fourth' First'