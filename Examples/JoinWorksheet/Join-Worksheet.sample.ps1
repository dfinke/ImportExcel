try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

#Create simple pages for 3 stores with product ID, Product Name, quanity price and total

@"
ID,Product,Quantity,Price,Total
12001,Nails,37,3.99,147.63
12002,Hammer,5,12.10,60.5
12003,Saw,12,15.37,184.44
12010,Drill,20,8,160
12011,Crowbar,7,23.48,164.36
"@ | ConvertFrom-Csv| Export-Excel -Path $xlSourcefile -WorkSheetname Oxford

@"
ID,Product,Quantity,Price,Total
12001,Nails,53,3.99,211.47
12002,Hammer,6,12.10,72.60
12003,Saw,10,15.37,153.70
12010,Drill,10,8,80
12012,Pliers,2,14.99,29.98
"@ | ConvertFrom-Csv| Export-Excel -Path $xlSourcefile -WorkSheetname Abingdon


@"
ID,Product,Quantity,Price,Total
12001,Nails,20,3.99,79.80
12002,Hammer,2,12.10,24.20
12010,Drill,11,8,88
12012,Pliers,3,14.99,44.97
"@ | ConvertFrom-Csv| Export-Excel -Path $xlSourcefile -WorkSheetname Banbury

#define a pivot table with a chart to show a sales by store, broken down by product
$ptdef = New-PivotTableDefinition -PivotTableName "Summary" -PivotRows "Store" -PivotColumns "Product" -PivotData @{"Total"="SUM"} -IncludePivotChart -ChartTitle "Sales Breakdown" -ChartType ColumnStacked -ChartColumn 10

#Join the 3 worksheets.
#Name the combined page "Total" and Name the column with the sheet names "store" (as the sheets 'Oxford','Abingdon' and 'Banbury' are the names of the stores
#Format the data as a table named "Summary", using the style "Light1", put the column headers in bold
#Put in a title and freeze to top of the sheet including title and colmun headings
#Add the Pivot table.
#Show the result
Join-Worksheet -Path $xlSourcefile -WorkSheetName "Total" -Clearsheet -FromLabel "Store" -TableName "Combined" -TableStyle Light1 -AutoSize -BoldTopRow -FreezePane 2,1 -Title "Store Sales Summary" -TitleBold -TitleSize 14  -PivotTableDefinition $ptdef -show
