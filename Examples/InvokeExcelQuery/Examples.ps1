try { Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 } catch { throw ; return }

$queries = 
'select * from [sheet1$A:A]',
'select * from [sheet1$]',
'select * from [sheet1$A2:E11]',
'select F2,F5 from [sheet1$A2:E11]',
'select * from [sheet1$A2:E11] where F2 = "Grocery"',
'select F2 as [Category], F5 as [Discount], F5*2 as [DiscountPlus] from [sheet1$A2:E11]'

foreach ($query in $queries) {
  "query: $($query)"
  Invoke-ExcelQuery .\testOleDb.xlsx $query | Format-Table  
}
