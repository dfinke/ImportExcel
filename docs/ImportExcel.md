---
Module Name: ImportExcelModule
Guid: 60dd4136-feff-401a-ba27-a84458c57ede
Download Help Link: https://dfinke.github.io/ImportExcel/Help
Version: 7.8.6.0
Locale: en-US
---

# ImportExcel Module

## Description

Automate Excel with PowerShell without having Excel installed. Works on Windows, Linux and Mac. Creating Tables, Pivot Tables, Charts and much more just got a lot easier.

## ImportExcel Cmdlets

### [Add-ConditionalFormatting](commands/Add-ConditionalFormatting.md)

Adds conditional formatting to all or part of a worksheet.

### [Add-ExcelChart](commands/Add-ExcelChart.md)

Creates a chart in an existing Excel worksheet.

### [Add-ExcelDataValidationRule](commands/Add-ExcelDataValidationRule.md)

Adds data validation to a range of cells

### [Add-ExcelName](commands/Add-ExcelName.md)

Adds a named-range to an existing Excel worksheet.

### [Add-ExcelTable](commands/Add-ExcelTable.md)

Adds Tables to Excel workbooks.

### [Add-PivotTable](commands/Add-PivotTable.md)

Adds a PivotTable (and optional PivotChart) to a workbook.

### [Add-Worksheet](commands/Add-Worksheet.md)

Adds a worksheet to an existing workbook.

### [BarChart](commands/BarChart.md)

```plaintext
BarChart [[-targetData] <Object>] [[-title] <Object>] [[-ChartType] <eChartType>] [-NoLegend] [-ShowCategory] [-ShowPercent] [<CommonParameters>]
```

### [Close-ExcelPackage](commands/Close-ExcelPackage.md)

Closes an Excel Package, saving, saving under a new name or abandoning changes and opening the file in Excel as required.

### [ColumnChart](commands/ColumnChart.md)

```plaintext
ColumnChart [[-targetData] <Object>] [[-title] <Object>] [[-ChartType] <eChartType>] [-NoLegend] [-ShowCategory] [-ShowPercent] [<CommonParameters>]
```

### [Compare-Worksheet](commands/Compare-Worksheet.md)

Compares two worksheets and shows the differences.

### [Convert-ExcelRangeToImage](commands/Convert-ExcelRangeToImage.md)

Gets the specified part of an Excel file and exports it as an image

### [ConvertFrom-ExcelData](commands/ConvertFrom-ExcelData.md)

{{ Fill in the Synopsis }}

### [ConvertFrom-ExcelSheet](commands/ConvertFrom-ExcelSheet.md)

Exports Sheets from Excel Workbooks to CSV files.

### [ConvertFrom-ExcelToSQLInsert](commands/ConvertFrom-ExcelToSQLInsert.md)

Generate SQL insert statements from Excel spreadsheet.

### [ConvertTo-ExcelXlsx](commands/ConvertTo-ExcelXlsx.md)

{{ Fill in the Synopsis }}

### [Copy-ExcelWorksheet](commands/Copy-ExcelWorksheet.md)

Copies a worksheet between workbooks or within the same workbook.

### [DoChart](commands/DoChart.md)

```plaintext
DoChart [[-targetData] <Object>] [[-title] <Object>] [[-ChartType] <eChartType>] [-NoLegend] [-ShowCategory] [-ShowPercent]
```

### [Enable-ExcelAutoFilter](commands/Enable-ExcelAutoFilter.md)

Enable the Excel AutoFilter

### [Enable-ExcelAutofit](commands/Enable-ExcelAutofit.md)

Make all text fit the cells

### [Expand-NumberFormat](commands/Expand-NumberFormat.md)

Converts short names for number formats to the formatting strings used in Excel

### [Export-Excel](commands/Export-Excel.md)

Exports data to an Excel worksheet.

### [Get-ExcelColumnName](commands/Get-ExcelColumnName.md)

{{ Fill in the Synopsis }}

### [Get-ExcelFileSchema](commands/Get-ExcelFileSchema.md)

Gets the schema of an Excel file.

### [Get-ExcelFileSummary](commands/Get-ExcelFileSummary.md)

Gets summary information on an Excel file like number of rows, columns, and more

### [Get-ExcelSheetDimensionAddress](commands/Get-ExcelSheetDimensionAddress.md)

Get the Excel address of the dimension of a sheet

### [Get-ExcelSheetInfo](commands/Get-ExcelSheetInfo.md)

Get worksheet names and their indices of an Excel workbook.

### [Get-ExcelWorkbookInfo](commands/Get-ExcelWorkbookInfo.md)

Retrieve information of an Excel workbook.

### [Get-HtmlTable](commands/Get-HtmlTable.md)

{{ Fill in the Synopsis }}

### [Get-Range](commands/Get-Range.md)

{{ Fill in the Synopsis }}

### [Get-XYRange](commands/Get-XYRange.md)

{{ Fill in the Synopsis }}

### [Import-Excel](commands/Import-Excel.md)

```plaintext
Import-Excel [-Path] <string[]> [[-WorksheetName] <string[]>] -NoHeader [-StartRow <int>] [-EndRow <int>] [-StartColumn <int>] [-EndColumn <int>] [-DataOnly] [-AsText <string[]>] [-AsDate <string[]>] [-Password <string>] [-ImportColumns <int[]>] [-Raw] [<CommonParameters>]

Import-Excel [-Path] <string[]> [[-WorksheetName] <string[]>] -HeaderName <string[]> [-StartRow <int>] [-EndRow <int>] [-StartColumn <int>] [-EndColumn <int>] [-DataOnly] [-AsText <string[]>] [-AsDate <string[]>] [-Password <string>] [-ImportColumns <int[]>] [-Raw] [<CommonParameters>]

Import-Excel [-Path] <string[]> [[-WorksheetName] <string[]>] [-StartRow <int>] [-EndRow <int>] [-StartColumn <int>] [-EndColumn <int>] [-DataOnly] [-AsText <string[]>] [-AsDate <string[]>] [-Password <string>] [-ImportColumns <int[]>] [-Raw] [<CommonParameters>]

Import-Excel [[-WorksheetName] <string[]>] -ExcelPackage <ExcelPackage> -NoHeader [-StartRow <int>] [-EndRow <int>] [-StartColumn <int>] [-EndColumn <int>] [-DataOnly] [-AsText <string[]>] [-AsDate <string[]>] [-Password <string>] [-ImportColumns <int[]>] [-Raw] [<CommonParameters>]

Import-Excel [[-WorksheetName] <string[]>] -ExcelPackage <ExcelPackage> -HeaderName <string[]> [-StartRow <int>] [-EndRow <int>] [-StartColumn <int>] [-EndColumn <int>] [-DataOnly] [-AsText <string[]>] [-AsDate <string[]>] [-Password <string>] [-ImportColumns <int[]>] [-Raw] [<CommonParameters>]

Import-Excel [[-WorksheetName] <string[]>] -ExcelPackage <ExcelPackage> [-StartRow <int>] [-EndRow <int>] [-StartColumn <int>] [-EndColumn <int>] [-DataOnly] [-AsText <string[]>] [-AsDate <string[]>] [-Password <string>] [-ImportColumns <int[]>] [-Raw] [<CommonParameters>]
```

### [Import-Html](commands/Import-Html.md)

{{ Fill in the Synopsis }}

### [Import-UPS](commands/Import-UPS.md)

{{ Fill in the Synopsis }}

### [Import-USPS](commands/Import-USPS.md)

{{ Fill in the Synopsis }}

### [Invoke-ExcelQuery](commands/Invoke-ExcelQuery.md)

Helper method for executing Read-OleDbData with some basic defaults.For additional help, see documentation for Read-OleDbData cmdlet.

### [Invoke-Sum](commands/Invoke-Sum.md)

{{ Fill in the Synopsis }}

### [Join-Worksheet](commands/Join-Worksheet.md)

Combines data on all the sheets in an Excel worksheet onto a single sheet.

### [LineChart](commands/LineChart.md)

```plaintext
LineChart [[-targetData] <Object>] [[-title] <Object>] [[-ChartType] <eChartType>] [-NoLegend] [-ShowCategory] [-ShowPercent] [<CommonParameters>]
```

### [Merge-MultipleSheets](commands/Merge-MultipleSheets.md)

Merges Worksheets into a single Worksheet with differences marked up.

### [Merge-Worksheet](commands/Merge-Worksheet.md)

Merges two Worksheets (or other objects) into a single Worksheet with differences marked up.

### [New-ConditionalFormattingIconSet](commands/New-ConditionalFormattingIconSet.md)

Creates an object which describes a conditional formatting rule a for 3,4 or 5 icon set.

### [New-ConditionalText](commands/New-ConditionalText.md)

Creates an object which describes a conditional formatting rule for single valued rules.

### [New-ExcelChartDefinition](commands/New-ExcelChartDefinition.md)

Creates a Definition of a chart which can be added using Export-Excel, or Add-PivotTable

### [New-ExcelStyle](commands/New-ExcelStyle.md)

{{ Fill in the Synopsis }}

### [New-PivotTableDefinition](commands/New-PivotTableDefinition.md)

Creates PivotTable definitons for Export-Excel

### [New-Plot](commands/New-Plot.md)

{{ Fill in the Synopsis }}

### [New-PSItem](commands/New-PSItem.md)

{{ Fill in the Synopsis }}

### [Open-ExcelPackage](commands/Open-ExcelPackage.md)

Returns an ExcelPackage object for the specified XLSX file.

### [PieChart](commands/PieChart.md)

```plaintext
PieChart [[-targetData] <Object>] [[-title] <Object>] [[-ChartType] <eChartType>] [-NoLegend] [-ShowCategory] [-ShowPercent] [<CommonParameters>]
```

### [Pivot](commands/Pivot.md)

```plaintext
Pivot [[-targetData] <Object>] [[-PivotRows] <Object>] [[-PivotData] <Object>] [[-ChartType] <eChartType>] [<CommonParameters>]
```

### [Read-Clipboard](commands/Read-Clipboard.md)

Read text from clipboard and pass to either ConvertFrom-Csv or ConvertFrom-Json.[Check out the how to video](https://youtu.be/dv2GOH5sbpA)

### [ReadClipboardImpl](commands/ReadClipboardImpl.md)

```plaintext
ReadClipboardImpl [-data] <string> [[-Delimiter] <Object>] [[-Header] <Object>] [<CommonParameters>]
```

### [Read-OleDbData](commands/Read-OleDbData.md)

Read data from an OleDb source using dotnet classes. This allows for OleDb queries against excel spreadsheets. Examples will only be for querying xlsx files.For additional documentation, see Microsoft's documentation on the System.Data OleDb namespace here:[System.Data.OleDb](https://docs.microsoft.com/en-us/dotnet/api/system.data.oledb)

### [Remove-Worksheet](commands/Remove-Worksheet.md)

Removes one or more worksheets from one or more workbooks

### [Select-Worksheet](commands/Select-Worksheet.md)

Sets the selected tab in an Excel workbook to be the chosen sheet and unselects all the others.

### [Send-SQLDataToExcel](commands/Send-SQLDataToExcel.md)

Inserts a DataTable - returned by a SQL query - into an ExcelSheet

### [Set-CellComment](commands/Set-CellComment.md)

{{ Fill in the Synopsis }}

### [Set-CellStyle](commands/Set-CellStyle.md)

{{ Fill in the Synopsis }}

### [Set-ExcelColumn](commands/Set-ExcelColumn.md)

Adds or modifies a column in an Excel worksheet, filling values, setting formatting and/or creating named ranges.

### [Set-ExcelRange](commands/Set-ExcelRange.md)

Applies number, font, alignment and/or color formatting, values or formulas to a range of Excel cells.

### [Set-ExcelRow](commands/Set-ExcelRow.md)

Fills values into a [new] row in an Excel spreadsheet, and sets row formats.

### [Set-WorksheetProtection](commands/Set-WorksheetProtection.md)

{{ Fill in the Synopsis }}

### [Test-Boolean](commands/Test-Boolean.md)

{{ Fill in the Synopsis }}

### [Test-Date](commands/Test-Date.md)

{{ Fill in the Synopsis }}

### [Test-Integer](commands/Test-Integer.md)

{{ Fill in the Synopsis }}

### [Test-Number](commands/Test-Number.md)

{{ Fill in the Synopsis }}

### [Test-String](commands/Test-String.md)

{{ Fill in the Synopsis }}

### [Update-FirstObjectProperties](commands/Update-FirstObjectProperties.md)

Updates the first object to contain all the properties of the object with the most properties in the array.