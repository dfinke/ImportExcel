# Overview of Module Functions

When available, Help page for function is linked.


| Function                          | Synopsis           | Description |
|-----------------------------------|-----------------------|-----------------------|
|[Add-ConditionalFormatting](mdHelp/en/add-conditionalformatting.md) |Adds conditional formatting to all or part of a worksheet.| Mark cells with icons, show a databar, change the color, font or format of cells with conditional formatting.|
|[Add-ExcelChart](mdHelp/en/add-excelchart.md)|Creates a chart in an existing Excel worksheet.|Create chart and optionally configure the type of chart, the range of X and Y value labels, the title, the legend, the ranges for both axes, the format and position of the axes.|
|[Add-ExcelDataValidationRule](mdHelp/en/add-exceldatavalidationrule.md)|Adds data validation to a range of cells|Excel supports the validation of user input. Ranges of cells can be marked to only contain numbers, dates, text up to a particular length, or selections from a list.|
|[Add-ExcelName](mdHelp/en/add-excelname.md)|Adds a named-range to an existing Excel worksheet.|It is often helpful to be able to refer to sets of cells with a name rather than using their co-ordinates; Add-ExcelName sets up these names.|
|[Add-ExcelTable](/mdHelp/en/add-exceltable.md)|Adds Tables to Excel workbooks.|Add table with unique name to workbook. Configure filter, header, totals, first and last column highlights.|
|[Add-PivotTable](mdHelp/en/add-pivottable.md)|Adds a PivotTable (and optional PivotChart) to a workbook.|If the PivotTable already exists, the source data will be updated.|
|[Add-Worksheet](mdHelp/en/add-worksheet.md)|Adds a worksheet to an existing workbook.|Worksheet added. Placement optionally configured.|
|BarChart|Creates bar chart in Excel worksheet.| Represents data with horizontal rectangular bars proportional to values represented. Can specify data range, axis labels, titles, colors, etc. |
|[Close-ExcelPackage](mdHelp/en/close-excelpackage.md)|Closes an Excel Package, saving, saving under a new name or abandoning changes and opening the file in Excel as required.|When working with an ExcelPackage object, the Workbook is held in memory and not saved until the .Save() method of the package is called. Close-ExcelPackage saves and disposes of the Package object.|
|ColumnChart|Creates column chart in Excel worksheet.| Represents data with vertical rectangular bars proportional to values represented. Can specify data range, axis labels, titles, colors, etc. |
|[Compare-Worksheet](mdHelp/en/compare-worksheet.md)|Compares two worksheets and shows the differences.|Reads the worksheet from each file and decides the column names and builds a hashtable of the key-column values and the rows in which they appear. It then uses PowerShell's Compare-Object command to compare the sheets (explicitly checking all the column names which have not been excluded).|
|[Convert-ExcelRangeToImage](mdHelp/en/convert-excelrangetoimage.md)|Gets the specified part of an Excel file and exports it as an image|Exports file range as image. Unlike most functions in the module it needs a local copy of Excel to be installed.|
|ConvertFrom-ExcelData|Converts data from Excel file into PowerShell objects.|Reads contents of Excel file and converts them into a collection of PowerShell objects.|
|[ConvertFrom-ExcelSheet](mdHelp/en/convertfrom-excelsheet.md)|Exports Sheets from Excel Workbooks to CSV files.|This command provides a convenient way to run Import-Excel @ImportParameters | Select-Object @selectParameters | export-Csv @ ExportParameters|
|[ConvertFrom-ExcelToSQLInsert](mdHelp/en/convertfrom-exceltosqlinsert.md)|Generate SQL insert statements from Excel spreadsheet.|Generate SQL insert statements from Excel spreadsheet.|
|ConvertTo-ExcelXlsx|Converts PowerShell objects or data tables into Excel file in xlsx format.| Allows export of data from PowerShell to an Excel file.|
|[Copy-ExcelWorksheet](mdHelp/en/copy-excelworksheet.md)|Copies a worksheet between workbooks or within the same workbook.|Copy-ExcelWorkSheet takes a Source object. The Destination workbook can be given as the path to an XLSx file, an ExcelPackage object or an ExcelWorkbook object.|
|DoChart| Creates chart of specified type in workbook.|Creates chart of specified type in workbook.|
|Enable-ExcelAutoFilter| Enables auto-filter feature for columns in Excel worksheet.| Allows user(s) to filter data based on specific criteria.|
|Enable-ExcelAutofit|Automatically adjusts width of columns in Excel worksheet to fit contents.|Ensures all data in columns is visible without truncation by dynamically adjusting column based on content.|
|[Expand-NumberFormat](mdHelp/en/expand-numberformat.md)|Converts short names for number formats, for example, 'Short-Date', to the formatting strings used in Excel|Converts short names for number formats to the formatting strings used in Excel|
|[Export-Excel](mdHelp/en/export-excel.md)|Exports data to an Excel worksheet.|Exports data to an Excel file and where possible tries to convert numbers in text fields so Excel recognizes them as numbers instead of text.|
|Get-ExcelColumnName|Returns column name(s) corresponding to specific column index or indices.|Returns column name(s) corresponding to specific column index or indices.|
|Get-ExcelFileSchema|Retrieves schema of Excel file.| Provides insight into Excel file structure.|
|Get-ExcelFileSummary| Retrieves summary of content and properties of Excel file.| Retrieves information such as number of worksheets, total rows, total columns, file size, and other metadata.|
|Get-ExcelSheetDimensionAddress| Retrieves address of the used range of worksheet in Excel file.| This identifies the range of cells within a worksheet that contain data or formatting. |
|[Get-ExcelSheetInfo](mdHelp/en/get-excelsheetinfo.md)|Get worksheet names and their indices of an Excel workbook.|The Get-ExcelSheetInfo cmdlet gets worksheet names and their indices of an Excel workbook.|
|[Get-ExcelWorkbookInfo](mdHelp/en/get-excelworkbookinfo.md)|Retrieve information of an Excel workbook.|The Get-ExcelWorkbookInfo cmdlet retrieves information (LastModifiedBy, LastPrinted, Created, Modified, ...) fron an Excel workbook. These are the same details that are visible in Windows Explorer when right clicking the Excel file, selecting Properties and check the Details tabpage.|
|Get-HtmlTable| Retrieves data from HTML table and converts it to PowerShell objects.|Retrieves data from HTML table and converts it to PowerShell objects.|
|Get-Range| Retrieves data within a specified range of cells in Excel worksheet.|Allows extraction of subset of data from worksheet using a range.|
|Get-XYRange| Retrieves data within a specified range of cells in an X-Y range.| Allows extraction of data from subset of rows and columns from worksheet using X/Y range.|
|[Import-Excel](mdHelp/en/import-excel.md)|Create custom objects from the rows in an Excel worksheet.|The Import-Excel cmdlet creates custom objects from the rows in an Excel worksheet. Each row is represented as one object.|
|Import-Html| Imports data from HTML file or URL into PowerShell.| Imports data from HTML file or URL into PowerShell.|
|Import-UPS| Imports UPS data into PowerShell.|Imports UPS data into PowerShell.|
|Import-USPS|Imports USPS data into PowerShell.|Imports USPS data into PowerShell.|
|Invoke-ExcelQuery| Executes query against Excel file and returns result as PowerShell objects.| Perform SQL-like queries on Excel data.|
|Invoke-Sum|Calculates sum of values in specified range of cells in an Excel worksheet.| Facilitates summation operations on data within Excel.|
|[Join-Worksheet](mdHelp/en/join-worksheet.md)|Combines data on all the sheets in an Excel worksheet onto a single sheet.|Join-Worksheet can work in two main ways, either Combining data which has the same layout from many pages into one, or Combining pages which have nothing in common.|
|LineChart|Creates line chart in Excel worksheet.| Represents data with series of data points connected by straight lines. |
|[Merge-MultipleSheets](mdHelp/en/merge-multiplesheets.md)|Merges Worksheets into a single Worksheet with differences marked up.|The Merge Worksheet command combines two sheets. Merge-MultipleSheets is designed to merge more than two.|
|[Merge-Worksheet](mdHelp/en/merge-worksheet.md)|Merges two Worksheets (or other objects) into a single Worksheet with differences marked up.|The Compare-Worksheet command takes two Worksheets and marks differences in the source document, and optionally outputs a grid showing the changes.|
|[New-ConditionalFormattingIconSet](mdHelp/en/new-conditionalformattingiconset.md)|Creates an object which describes a conditional formatting rule a for 3,4 or 5 icon set.|This command builds the defintion of a Conditional formatting rule for an icon set.|
|[New-ConditionalText](mdHelp/en/new-conditionaltext.md)|Creates an object which describes a conditional formatting rule for single valued rules.|Some Conditional formatting rules don't apply styles to a cell (IconSets and Databars); some take two parameters (Between); some take none (ThisWeek, ContainsErrors, AboveAverage etc).The others take a single parameter (Top, BottomPercent, GreaterThan, Contains etc).This command creates an object to describe the last two categories, which can then be passed to Export-Excel.|
|[New-ExcelChartDefinition](mdHelp/en/new-excelchartdefinition.md)|Creates a Definition of a chart which can be added using Export-Excel, or Add-PivotTable.|All the parameters which are passed to Add-ExcelChart can be added to a chart-definition object and passed to Export-Excel with the -ExcelChartDefinition parameter, or to Add-PivotTable with the -PivotChartDefinition parameter. This command sets up those definition objects.|
|New-ExcelStyle| Creates a new style object that defines various formatting properties, such as font size, font color, background color, borders, and alignment.| After being created, style object can be applied to specific cells or ranges.|
|[New-PivotTableDefinition](mdHelp/en/new-pivottabledefinition.md)|Creates PivotTable definitons for Export-Excel|Export-Excel allows a single PivotTable to be defined using the parameters -IncludePivotTable, -PivotColumns, -PivotRows, -PivotData, -PivotFilter, -PivotTotals, -PivotDataToColumn, -IncludePivotChart and -ChartType. Its -PivotTableDefintion paramater allows multiple PivotTables to be defined, with additional parameters. New-PivotTableDefinition is a convenient way to build these definitions.|
|New-Plot| Creates new plot in Excel worksheet. |Creates new plot in Excel worksheet. |
|New-PSItem| Creates a new PSItem.|Creates a new PSItem.|
|[Open-ExcelPackage](mdHelp/en/open-excelpackage.md)|Returns an ExcelPackage object for the specified XLSX file.|Import-Excel and Export-Excel open an Excel file, carry out their tasks and close it again. Sometimes it is necessary to open a file and do other work on it. Open-ExcelPackage allows the file to be opened for these tasks.|
|PieChart| Creates pie chart in Excel worksheet.| Represents data as circular graph divided into slices. |
|Pivot|Creates pivot table in Excel worksheet.| Pivot tables unlock powerful data analysis, especially on large data sets. Can group and aggregate data, apply filters and generate summary stats. |
|Read-Clipboard|Reads data from clipboard and imports it into PowerShell.|Leverage data copied to the clipboard from an external source in PowerShell script.|
|Read-OleDbData|Reads data from a data source using OleDb (Object Linking and Embedding Database) connectivity.|Allows query of data from various data sources such as Excel files, Access databases, or other databases that support OleDb connections.|
|ReadClipboardImpl|Internal implementation detail for Read-Clipboard.|Internal implementation detail for Read-Clipboard.|
|[Remove-Worksheet](mdHelp/en/remove-worksheet.md)|Removes one or more worksheets from one or more workbooks|Removes one or more worksheets from one or more workbooks.|
|[Select-Worksheet](mdHelp/en/select-worksheet.md)|Sets the selected tab in an Excel workbook to be the chosen sheet and unselects all the others.|Sometimes when a sheet is added we want it to be the active sheet, sometimes we want the active sheet to be left as it was. Select-Worksheet exists to change which sheet is the selected tab when Excel opens the file.|
|[Send-SQLDataToExcel](mdHelp/en/send-sqldatatoexcel.md)|Inserts a DataTable - returned by a SQL query - into an ExcelSheet| This command takes a SQL statement and run it against a database connection. After fetching the data it calls Export-Excel with the data as the value of InputParameter and whichever of Export-Excel's parameters it was passed|
|Set-CellComment|Adds or modifies comments in specific cells.| Used to add additional context or explanations for data in workbook.|
|Set-CellStyle|Sets style of one or multiple cells in a worksheet.| Customize font, font size, font color, fill color, border and alignment of cells.|
|[Set-ExcelColumn](mdHelp/en/set-excelcolumn.md)|Adds or modifies a column in an Excel worksheet, filling values, setting formatting and/or creating named ranges.|Set-ExcelColumn can take a value which is either a string containing a value or formula or a scriptblock which evaluates to a string, and optionally a column number and fills that value down the column. A column heading can be specified, and the column can be made a named range. The column can be formatted in the same operation.|
|[Set-ExcelRange](mdHelp/en/set-excelrange.md)|Applies number, font, alignment and/or color formatting, values or formulas to a range of Excel cells.| Style elements for a range of cells, this includes auto-sizing and hiding, setting font elements (Name, Size, Bold, Italic, Underline & UnderlineStyle and Subscript & SuperScript), font and background colors, borders, text wrapping, rotation, alignment within cells, and number format.|
|[Set-ExcelRow](mdHelp/en/set-excelrow.md)|Fills values into a new row in an Excel spreadsheet, and sets row formats.|Fills values into a [new] row in an Excel spreadsheet, and sets row formats.|
|Set-WorksheetProtection| Protects a worksheet within an Excel workbook by applying various protection settings.| Can include password protection, locking cells, hiding formulas, restricting formatting, and disallowing certain row/column operations.|
|Test-Boolean| Test whether a given value is a boolean (true/false) or not.| Return true if it is and false otherwise.
|Test-Date| Test whether a given value is a valid date or not.| Return true if it is and false otherwise.
|Test-Integer| Test whether a given value is an integer or not.| Return true if it is and false otherwise.
|Test-Number| Test whether a given value is a numeric type or not. | Return true if it is and false otherwise.
|Test-String| Test whether a given value is a string or not.| Return true if it is and false otherwise.
|[Update-FirstObjectProperties](mdHelp/en/update-firstobjectproperties.md)|Updates the first object to contain all the properties of the object with the most properties in the array.| This is usefull when not all objects have the same quantity of properties and CmdLets like Out-GridView or Export-Excel are not able to show all the properties because the first object doesn't have them all.|
|Convert-XlRangeToImage| Function converts a range of cells in an Excel worksheet to an image file.|Takes a specified range of cells as input and generates an image file (such as PNG or JPEG) containing the visual representation of the cells.This allows the export a portion of an Excel worksheet as an image, for example, to embed it in a document or share it in a presentation.|
|Export-ExcelSheet| Exports the content of an Excel worksheet to a new Excel workbook or to another format, such as CSV or HTML.|Allows specification of the output path and format for the exported data. Can export the entire worksheet or specify a range of cells to export.
|New-ExcelChart| Creates a new chart of specified type in an Excel worksheet based on the specified data range.|It allows specification the type of chart (e.g., bar chart, line chart, pie chart) and customize various aspects of the chart's appearance, such as title, axis labels, and data series. Once created, the chart can be inserted into the worksheet at a specified location.|
|Set-Column|Same as Set-ExcelColumn|Same as Set-ExcelColumn|
|Set-Format|Apply formatting to cells within a specified range.| Define various formatting options such as font size, font color, and cell background color.
|Set-Row|Same as Set-ExcelRow|Same as Set-ExcelRow|
|Use-ExcelData|Allows utilization of Excel data directly within PowerShell script.|Simplifies the process of working with Excel data in PowerShell scripts, providing a convenient and efficient way to leverage Excel data for various tasks.



