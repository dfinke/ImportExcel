**ImportExcel**, a [PowerShell Module](https://docs.microsoft.com/en-us/powershell/scripting/developer/module/understanding-a-windows-powershell-module), allows you to read and write Excel files without installing Microsoft Excel on your system. No need to bother with the cumbersome Excel COM-object. Creating Tables, Pivot Tables, Charts and much more has just become a lot easier.

- [Video Demonstration](https://www.youtube.com/watch?v=fvKKdIzJCws&list=PL5uoqS92stXioZw-u-ze_NtvSo0k0K0kq)

*If this project helps you reduce the time to get your job done, let me know!*

<p>
<a alt="PayPal" href="https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=UCSB9YVPFSNCY"><img src="https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif"></a>
<img src="https://media.giphy.com/media/hpXxJ78YtpT0s/giphy.gif">
</p>

<!-- BADGES -->

<p>
<a href="https://www.powershellgallery.com/packages/ImportExcel"><img src="https://img.shields.io/powershellgallery/v/ImportExcel.svg"></a>
<a href="https://www.powershellgallery.com/packages/ImportExcel"><img src="https://img.shields.io/powershellgallery/dt/ImportExcel.svg"></a>
<a href="./LICENSE.txt"><img src="https://img.shields.io/badge/License-Apache%202.0-blue.svg"></a>
</p>

| CI System    | Environment                   | Status                                                                                                                                                                                                                                          |
|--------------|-------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Azure DevOps | Windows                       | [![Build Status](https://dougfinke.visualstudio.com/ImportExcel/_apis/build/status/dfinke.ImportExcel?branchName=master&jobName=Windows)](https://dougfinke.visualstudio.com/ImportExcel/_build/latest?definitionId=21&branchName=master)       |
| Azure DevOps | Windows (Core)                | [![Build Status](https://dougfinke.visualstudio.com/ImportExcel/_apis/build/status/dfinke.ImportExcel?branchName=master&jobName=WindowsPSCore)](https://dougfinke.visualstudio.com/ImportExcel/_build/latest?definitionId=21&branchName=master) |
| Azure DevOps | Ubuntu                        | [![Build Status](https://dougfinke.visualstudio.com/ImportExcel/_apis/build/status/dfinke.ImportExcel?branchName=master&jobName=Ubuntu)](https://dougfinke.visualstudio.com/ImportExcel/_build/latest?definitionId=21&branchName=master)        |
| Azure DevOps | macOS                         | [![Build Status](https://dougfinke.visualstudio.com/ImportExcel/_apis/build/status/dfinke.ImportExcel?branchName=master&jobName=macOS)](https://dougfinke.visualstudio.com/ImportExcel/_build/latest?definitionId=21&branchName=master)         |

<!-- /BADGES -->



# Installation


### [PowerShell v7.x](https://github.com/PowerShell/Powershell) / [PowerShell v5.1](https://www.microsoft.com/en-us/download/details.aspx?id=50395)

You can install the `ImportExcel` module directly from the PowerShell Gallery

- {Recommended} **Install for the current user:**
```PowerShell
Install-Module ImportExcel -scope CurrentUser
```
- {Requires Elevation} **Install for all users:**
```PowerShell
Install-Module ImportExcel -scope AllUsers
```



# What's New in ImportExcel 7.1.1


- Merged [Nate Ferrell](https://github.com/scrthq)'s Linux fix. Thanks!
- Moved `Export-MultipleExcelSheets` from psm1 to Examples/Experimental
- Deleted the CI build in Appveyor
- Thank you [Mikey Bronowski](https://github.com/MikeyBronowski) for 
    - Multiple sweeps 
    - Standardising casing of parameter names, and variables
    - Plus updating > 50 of the examples and making them consistent. 



> [*Past Release Notes*](CHANGELOG.md)




# Known Issues

* Using `-IncludePivotTable`, if that pivot table name exists, you'll get an error.
    * We're investigating a solution.
    * *Workaround* — Delete the Excel file first, then do the export.




# Function Overview

### Add-ConditionalFormatting 

Adds conditional formatting to all or part of a worksheet.

- [Learn more...](mdHelp/en/Add-ConditionalFormatting.md)


### Add-ExcelChart 

Creates a chart in an existing Excel worksheet.

- [Learn more...](mdHelp/en/Add-ExcelChart.md)


### Add-ExcelDataValidationRule 

Adds data validation to a range of cells

- [Learn more...](mdHelp/en/Add-ExcelDataValidationRule.md)


### Add-ExcelName 

Adds a named-range to an existing Excel worksheet.

- [Learn more...](mdHelp/en/Add-ExcelName.md)


### Add-ExcelTable 

Adds Tables to Excel workbooks.

- [Learn more...](mdHelp/en/Add-ExcelTable.md)


### Add-PivotTable 

Adds a PivotTable (and optional PivotChart) to a workbook.

- [Learn more...](mdHelp/en/Add-PivotTable.md)


### Add-WorkSheet 

Adds a worksheet to an existing workbook.

- [Learn more...](mdHelp/en/Add-WorkSheet.md)


### Close-ExcelPackage 

Closes an Excel Package, either saving normally, saving under a new name, or abandoning changes, in addition to opening the file in Excel (if specified).

- [Learn more...](mdHelp/en/Close-ExcelPackage.md)


### Compare-WorkSheet 

Compares two worksheets and shows the differences.

- [Learn more...](mdHelp/en/Compare-WorkSheet.md)


### Convert-ExcelRangeToImage 

Gets the specified part of an Excel file and exports it as an image

- [Learn more...](mdHelp/en/Convert-ExcelRangeToImage.md)


### ConvertFrom-ExcelSheet 

Exports Sheets from Excel Workbooks to CSV files.

- [Learn more...](mdHelp/en/ConvertFrom-ExcelSheet.md)


### ConvertFrom-ExcelToSQLInsert 

Generate SQL insert statements from Excel spreadsheet.

- [Learn more...](mdHelp/en/ConvertFrom-ExcelToSQLInsert.md)


### Copy-ExcelWorkSheet 

Copies a worksheet between workbooks or within the same workbook.

- [Learn more...](mdHelp/en/Copy-ExcelWorkSheet.md)


### Expand-NumberFormat 

Converts short names for number formats to the formatting strings used in Excel.

- [Learn more...](mdHelp/en/Expand-NumberFormat.md)


### Export-Excel 

Exports data to an Excel worksheet.

- [Learn more...](mdHelp/en/Export-Excel.md)


### Get-ExcelSheetInfo 

Get worksheet names and their indices of an Excel workbook.

- [Learn more...](mdHelp/en/Get-ExcelSheetInfo.md)


### Get-ExcelWorkbookInfo 

Retrieve information of an Excel workbook.

- [Learn more...](mdHelp/en/Get-ExcelWorkbookInfo.md)


### Import-Excel 

Create custom objects from the rows in an Excel worksheet.

- [Learn more...](mdHelp/en/Import-Excel.md)


### Join-Worksheet 

Combines data on all the sheets in an Excel worksheet onto a single sheet.

- [Learn more...](mdHelp/en/Join-Worksheet.md)


### Merge-MultipleSheets 

Merges Worksheets into a single Worksheet with differences marked up.

- [Learn more...](mdHelp/en/Merge-MultipleSheets.md)


### Merge-Worksheet 

Merges two Worksheets (or other objects) into a single Worksheet with differences marked up.

- [Learn more...](mdHelp/en/Merge-Worksheet.md)


### New-ConditionalFormattingIconSet 

Creates an object which describes a conditional formatting rule a for 3, 4, or 5 icon set.

- [Learn more...](mdHelp/en/New-ConditionalFormattingIconSet.md)


### New-ConditionalText 

Creates an object which describes a conditional formatting rule for single valued rules.

- [Learn more...](mdHelp/en/New-ConditionalText.md)


### New-ExcelChartDefinition 

Creates a Definition of a chart which can be added using Export-Excel or Add-PivotTable.

- [Learn more...](mdHelp/en/New-ExcelChartDefinition.md)


### New-PivotTableDefinition 

Creates PivotTable definitions for Export-Excel.

- [Learn more...](mdHelp/en/New-PivotTableDefinition.md)


### Open-ExcelPackage 

Returns an ExcelPackage object for the specified XLSX file.

- [Learn more...](mdHelp/en/Open-ExcelPackage.md)


### Remove-WorkSheet 

Removes one or more worksheets from one or more workbooks.

- [Learn more...](mdHelp/en/Remove-WorkSheet.md)


### Select-Worksheet 

Sets the selected tab in an Excel workbook to be the chosen sheet, deselecting all others.

- [Learn more...](mdHelp/en/Select-Worksheet.md)


### Send-SQLDataToExcel 

Inserts a DataTable (returned by a SQL query) into an ExcelSheet.

- [Learn more...](mdHelp/en/Send-SQLDataToExcel.md)


### Set-ExcelColumn 

Adds or modifies a column in an Excel worksheet, filling values, setting formatting, and/or creating named ranges.

- [Learn more...](mdHelp/en/Set-ExcelColumn.md)


### Set-ExcelRange 

Applies number, font, alignment, and/or color formatting values (or formulas) to a range of Excel cells.

- [Learn more...](mdHelp/en/Set-ExcelRange.md)


### Set-ExcelRow 

Fills values into a \[new\] row in an Excel spreadsheet and sets row formats.

- [Learn more...](mdHelp/en/Set-ExcelRow.md)


### Update-FirstObjectProperties 

Updates the first object to contain all the properties of the object with the most properties in the array.

- [Learn more...](mdHelp/en/Update-FirstObjectProperties.md)




# Examples

`gsv | Export-Excel .\test.xlsx -WorkSheetname Services`

`dir -file | Export-Excel .\test.xlsx -WorkSheetname Files`

`ps | Export-Excel .\test.xlsx -WorkSheetname Processes -IncludePivotTable -Show -PivotRows Company -PivotData PM`


### Convert (All or Some) Excel Sheets to Text files

Reads each sheet in TestSheets.xlsx and outputs it to the data directory as the sheet name with the extension .txt

    ConvertFrom-ExcelSheet .\TestSheets.xlsx .\data

Reads and outputs sheets like Sheet10 and Sheet20 form TestSheets.xlsx and outputs it to the data directory as the sheet name with the extension .txt

    ConvertFrom-ExcelSheet .\TestSheets.xlsx .\data sheet?0


### Adding a Title

You can set the pattern, size and of if the title is bold.

    $p=@{
        Title = "Process Report as of $(Get-Date)"
        TitleFillPattern = "LightTrellis"
        TitleSize = 18
        TitleBold = $true

        Path  = "$pwd\testExport.xlsx"
        Show = $true
        AutoSize = $true
    }

    Get-Process |
        Where Company | Select Company, PM |
        Export-Excel @p

![image](https://raw.githubusercontent.com/dfinke/ImportExcel/master/images/Title.png)


### Using Export-MultipleExcelSheets

![image](https://raw.githubusercontent.com/dfinke/ImportExcel/master/images/ExportMultiple.gif)

    $p = Get-Process

    $DataToGather = @{
        PM        = {$p|select company, pm}
        Handles   = {$p|select company, handles}
        Services  = {gsv}
        Files     = {dir -File}
        Albums    = {(Invoke-RestMethod http://www.dougfinke.com/PowerShellfordevelopers/albums.js)}
    }

    Export-MultipleExcelSheets -Show -AutoSize .\testExport.xlsx $DataToGather

***NOTE*** If the sheet exists when using *-WorkSheetname* parameter, it will be deleted and then added with the new data.


### Total Physical Memory Grouped By Company

![image](https://raw.githubusercontent.com/dfinke/ImportExcel/master/images/PivotTablesAndCharts.png)


### Importing data from an Excel spreadsheet

![image](https://raw.githubusercontent.com/dfinke/ImportExcel/master/images/TryImportExcel.gif)




# Testimonials

[![](Testimonials/testimonial01.png)](https://twitter.com/pacdelory/status/713791929327038464)

[![](Testimonials/testimonial02.png)](https://twitter.com/Tmjr75/status/991762683920830464)




# Credits

Big thanks to [Illy](https://github.com/ili101) for taking the Azure DevOps CI to the next level. Improved badges, improved matrix for cross platform OS testing and more.

Plus, wiring the [PowerShell ScriptAnalyzer Excel report](https://github.com/dfinke/ImportExcel/pull/590#issuecomment-488659081) we built into each run as an artifact.

![](images/ScriptAnalyzerReport.png)