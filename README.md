# PowerShell and Excel 

![](images/logoWithInstall.png)

<br/>

[![Build Status](https://dougfinke.visualstudio.com/ImportExcel/_apis/build/status/dfinke.ImportExcel?branchName=master)](https://dougfinke.visualstudio.com/ImportExcel/_build)
[![](https://img.shields.io/powershellgallery/v/ImportExcel.svg)](https://www.powershellgallery.com/packages/ImportExcel)
[![](https://img.shields.io/powershellgallery/dt/ImportExcel.svg)](https://www.powershellgallery.com/packages/ImportExcel)
[![](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://github.com/dfinke/ImportExcel/tree/70ab9e46c776e96fb287682d5b9b4b51a0ec3bac/LICENSE.txt)
<a href="https://www.paypal.com/paypalme/DougCharlesFinke"><img src="https://img.shields.io/badge/Donate-PayPal-green.svg" alt="Donate"></a>

# Overview

Automate Excel with PowerShell without having Excel installed. Works on Windows, Linux and MAC. Creating Tables, Pivot Tables, Charts and much more just got a lot easier.

## Examples 
Check out [more than 100 examples](Examples/) on ways to make you more productive.

# Basic Usage
## Installation

```powershell
Install-Module -Name ImportExcel
```

## Create a spreadsheet
Here is a quick example that will create spreadsheet file from CSV data. Works with JSON, Databases, and more.

```powershell
$data = ConvertFrom-Csv @"
Region,State,Units,Price
West,Texas,927,923.71
North,Tennessee,466,770.67
East,Florida,520,458.68
East,Maine,828,661.24
West,Virginia,465,053.58
North,Missouri,436,235.67
South,Kansas,214,992.47
North,North Dakota,789,640.72
South,Delaware,712,508.55
"@

$data | Export-Excel .\salesData.xlsx
```

![](images/salesdata.png)

## Read a spreadsheet

Quickly read a spreadsheet document into a PowerShell array.

```powershell
$data = Import-Excel .\salesData.xlsx

$data
```

```powershell
Region State        Units Price
------ -----        ----- -----
West   Texas        927   923.71
North  Tennessee    466   770.67
East   Florida      520   458.68
East   Maine        828   661.24
West   Virginia     465   053.58
North  Missouri     436   235.67
South  Kansas       214   992.47
North  North Dakota 789   640.72
South  Delaware     712   508.55
```

## Add a chart to spreadsheet

Chart generation is as easy as 123. Building charts based on data in your worksheet doesn't get any easier.

Plus, it is automated and repeatable.

```powershell
$data = ConvertFrom-Csv @"
Region,State,Units,Price
West,Texas,927,923.71
North,Tennessee,466,770.67
East,Florida,520,458.68
East,Maine,828,661.24
West,Virginia,465,053.58
North,Missouri,436,235.67
South,Kansas,214,992.47
North,North Dakota,789,640.72
South,Delaware,712,508.55
"@

$chart = New-ExcelChartDefinition -XRange State -YRange Units -Title "Units by State" -NoLegend

$data | Export-Excel .\salesData.xlsx -AutoNameRange -ExcelChartDefinition $chart -Show
```

![](images/SalesDataChart.png)

## Add a pivot table to spreadsheet

Categorize, sort, filter, and summarize any amount data with pivot tables. Then add charts.

```powershell
$data = ConvertFrom-Csv @"
Region,State,Units,Price
West,Texas,927,923.71
North,Tennessee,466,770.67
East,Florida,520,458.68
East,Maine,828,661.24
West,Virginia,465,053.58
North,Missouri,436,235.67
South,Kansas,214,992.47
North,North Dakota,789,640.72
South,Delaware,712,508.55
"@

$data | Export-Excel .\salesData.xlsx -AutoNameRange -Show -PivotRows Region -PivotData @{'Units'='sum'} -PivotChartType PieExploded3D
```

![](images/SalesDataChartPivotTable.png)

# Convert Excel data to other formats

## Create a separate CSV file for each Excel sheet

Do you have an Excel file with multiple sheets and you need to convert each sheet to CSV file?

### Problem Solved

The [yearlyRetailSales.xlsx](Examples/Import-Excel) has 12 sheets of retail data for the year.

This single line of PowerShell converts any number of sheets in an Excel workbook to separate CSV files.

```powershell
(Import-Excel .\yearlyRetailSales.xlsx *).GetEnumerator() |
ForEach-Object { $_.Value | Export-Csv ($_.key + '.csv') }
```

# Additional Resources

## Videos

- [Export-Excel Hello World](https://youtu.be/fvKKdIzJCws?list=PL5uoqS92stXioZw-u-ze_NtvSo0k0K0kq)
- [Make Excel Data Pop](https://youtu.be/gQaYI5hxqM4?list=PL5uoqS92stXioZw-u-ze_NtvSo0k0K0kq)
- [Slice And Dice Data](https://youtu.be/kzllxvqr3TY?list=PL5uoqS92stXioZw-u-ze_NtvSo0k0K0kq)
- [Lightning talk - PowerShell Excel Module](https://youtu.be/znVu2q11Rp4?list=PL5uoqS92stXioZw-u-ze_NtvSo0k0K0kq)

## More Videos

- [Look smarter: deliver your work in Excel](https://youtu.be/tu8Mfkwi8zI) - James O'Neill [@jamesoneill](https://twitter.com/jamesoneill)
- [Module Monday: ImportExcel](https://youtu.be/rBA_IeTmCb8?t=5) - Adam Driscoll [@adamdriscoll](https://twitter.com/adamdriscoll)
- [Tutorials Excel Module Part 1](https://youtu.be/2cwBuYbZ3To)
- [Tutorials Excel Module Part 2](https://youtu.be/8ojg-qjOnVI)
- [Tutorials Excel Module Part 3](https://youtu.be/3IgASPD0UrQ)
- [PowerShell Excel - Invoke-ExcelQuery](https://youtu.be/_7xuhsZm0Ao)
- [Powershell Excel - Data Validation](https://youtu.be/NGhahuY8j1M)
- [Creating Dashboards xPlatform](https://youtu.be/qMWkZt6ikgM)

## Articles

|Title|Author|Twitter|
|------|------|------|
|[More tricks with PowerShell and Excel](https://jamesone111.wordpress.com/2018/05/31/more-tricks-with-powershell-and-excel/)|James O'Neill|[@jamesoneill](https://twitter.com/jamesoneill)|
|[Using the Import-Excel module: Part 1 Importing](https://jamesone111.wordpress.com/2017/12/05/using-the-import-excel-part-1-importing/)|James O'Neill|[@jamesoneill](https://twitter.com/jamesoneill)|
|[Using the Import Excel module part 2: putting data into .XLSx files](https://jamesone111.wordpress.com/2017/12/11/using-the-import-excel-module-part-2-putting-data-into-xlsx-files/)|James O'Neill|[@jamesoneill](https://twitter.com/jamesoneill)|
|[Using the import Excel Module: Part 3, Pivots and charts, data and calculations](https://jamesone111.wordpress.com/2017/12/12/using-the-import-excel-module-part-3-pivots-and-charts-data-and-calculations/)|James O'Neill|[@jamesoneill](https://twitter.com/jamesoneill)|
|[Export AdventureWorksDW2017 to Excel for a Power BI Demo with Export-Excel](https://sqlvariant.com/2019/03/export-adventureworksdw2017-to-excel-for-a-powerbi-demo-with-export-excel-in-powershell/)|Aaron Nelson|[@sqlvariant](https://twitter.com/sqlvariant)
|[Creating beautiful Powershell Reports in Excel](https://dfinke.github.io/powershell/2019/07/31/Creating-beautiful-Powershell-Reports-in-Excel.html)|Doug Finke|[@dfinke](https://twitter.com/dfinke)
|[PowerShell Excel and Conditional Formatting](https://dfinke.github.io/powershell/2020/05/02/PowerShell-Excel-and-Conditional-Formatting.html)|Doug Finke|[@dfinke](https://twitter.com/dfinke)
|[Learn to Automate Excel like a Pro with PowerShell](https://dfinke.github.io/powershell/2019/08/29/Learn-to-Automate-Excel-like-a-Pro-with-PowerShell.html)|Doug Finke|[@dfinke](https://twitter.com/dfinke)

## Contributing
Contributions are welcome! Open a pull request to fix a bug, or open an issue to discuss a new feature or change.

Original [README.md](./README.original.md)