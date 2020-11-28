# Multiplecharts

## PowerShell

```text
$xlFile = "$env:TEMP\ImportExcelExample.xlsx"
Remove-Item $xlFile -ErrorAction Ignore

$data =  ConvertFrom-Csv @"
ID,Product,Quantity,Price,Total
12001,Nails,37,3.99,147.63
12002,Hammer,5,12.10,60.5
12003,Saw,12,15.37,184.44
12010,Drill,20,8,160
12011,Crowbar,7,23.48,164.36
"@

$chart1 = New-ExcelChartDefinition -YRange "Price"   -XRange "Product" -Title "Item price"   -NoLegend -Height 225
$chart2 = New-ExcelChartDefinition -YRange "Total   "-XRange "Product" -Title "Total sales"  -NoLegend -Height 225 -Row 9  -Column 15
$chart3 = New-ExcelChartDefinition -YRange "Quantity"-XRange "Product" -Title "Sales volume" -NoLegend -Height 225 -Row 15

$data | Export-Excel -Path $xlFile -AutoFilter -AutoNameRange -AutoSize -Show -ExcelChartDefinition $chart1,$chart2,$chart3
```

## Result

![](https://raw.githubusercontent.com/dfinke/ImportExcel/master/Examples/Charts/Multiplecharts.png)

