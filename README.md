# PowerShell and Excel 

![](images/logo.png)

# Overview

Automate Excel with PowerShell without having Excel installed. Works on Windows, Linux and MAC. Creating Tables, Pivot Tables, Charts and much more has just become a lot easier.

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

![](images/salesDataChart.png)

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

# Bonus Points

## Create a separate CSV file for each Excel sheet

Do you have a Excel file with multiple sheets and you need to convert each sheet to CSV file?

### Problem Solved

The `yearlyRetailSales.xlsx` has 12 sheets of retail data for the year.

This single line of PowerShell converts any number of sheets in an Excel workbook to a separate CSV file.

```powershell
(Import-Excel .\yearlyRetailSales.xlsx *).GetEnumerator() |
ForEach-Object { $_.Value | Export-Csv ($_.key + '.csv') }
```