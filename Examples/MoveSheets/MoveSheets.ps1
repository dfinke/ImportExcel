try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$xlfile = "$env:TEMP\testThis.xlsx"
Remove-Item $xlfile -ErrorAction Ignore

1..10  | Export-Excel $xlfile -WorkSheetname First
11..20 | Export-Excel $xlfile -WorkSheetname Second -MoveToStart
21..30 | Export-Excel $xlfile -WorkSheetname Third  -MoveBefore First
31..40 | Export-Excel $xlfile -WorkSheetname Fourth -MoveAfter Third -Show