# 7.6.0

- **_[Under investigation]_** Fix -StartRow and -StartColumn being ignored.
- James O'Neill: 
    - Update Get-HtmlTable to support to use PowerHTML (maintained by [Justin Grote](https://twitter.com/**JustinWGrote**)). 
    - Added example to including a new function Import-ByColumn. Works like Import-Excel but with data in columns instead of the conventional rows.
- Update Import-HTML with better defaults
- Fixed example `Get-ModuleStats.ps1` which reads the PowerShell Gallery page and extracts the stats table


# v7.5.2
- Changed the switch `-NotAsDictionary` to `-Raw`. Works with `-Worksheetname *` reads all the sheets in the xlsx file and returns an array.

# v7.5.1
- Fixed `Import-Excel` - Reset `EndRow` and `EndColumn` in the correct place.
# v7.5.0
## Fixes

- Importing multiple files with Import-Excel by pipeline uses only the first file for the row count https://github.com/dfinke/ImportExcel/issues/1172

## New Features

- Import-Excel now supports importing multiple sheets. It can either return a dictionary of all sheets, or as a single array of all sheets combined.
    - `Import-Excel $xlfile *`                  # reads all sheets, returns all data in a dictionary
    - `Import-Excel $xlfile * -NotAsDictionary` # reads all sheets, returns all data in a single array
- Added helper functions. Useful for working with an Excel package via `Open-ExcelPackage` or `-PassThru` 
    - `Enable-ExcelAutoFilter`
    - `Enable-ExcelAutofit`
    - `Get-ExcelSheetDimensionAddress`

# v7.4.2

- Thank you [James Mueller](https://github.com/jamesmmueller) Updated `ConvertFrom-ExcelToSQLInsert` to handle single quotes in the SQL statement.

- Thank you to Josh Hendricks
    - Add images to spreadsheets. [Check it out](https://github.com/dfinke/ImportExcel/tree/master/Examples/AddImage)
        - Catch up with him on [GitHub](https://github.com/joshooaj) and [Twitter](https://twitter.com/joshooaj) for the idea

# v7.4.1

- Implements: https://github.com/dfinke/ImportExcel/issues/1111
- Refactored ReZip into separate function
- Deletes temp folder after rezipping
- Added -ReZip to `Close-ExcelPackage`

# v7.4.0

- Thank you to [Max Goczall](https://github.com/muschebubusche) for this contribution!
    - `ImportColumns` parameter added to `ImportExcel`. It is used to define which columns of the ExcelPackage should be imported.

```powershell
Import-Excel -Path $xlFile -ImportColumns @(6,7,12,25,46)
```

# v7.3.1

- Added query Excel spreadsheets, with SQL queries!

```powershell
$query = 'select F2 as [Category], F5 as [Discount], F5*2 as [DiscountPlus] from [sheet1$A2:E11]'

Invoke-ExcelQuery .\testOleDb.xlsx $query
```

![](./images/SQL-Spreadsheet.png)

## Result

```
Category    Discount DiscountPlus
--------    -------- ------------
Cosmetics        0.7          1.4
Grocery          0.3          0.6
Apparels         0.2          0.4
Electronics      0.1          0.2
Electronics        0            0
Apparels         0.8          1.6
Electronics      0.7          1.4
Cosmetics        0.6          1.2
Grocery          0.4          0.8
Grocery          0.3          0.6
```

- Thank you to Roy Ashbrook for the SQL query code. Catch up with Roy:

|Media|Link|
|---|---|
|twitter|https://twitter.com/royashbrook
|github|https://github.com/royashbrook
|linkedin|https://linkedin.com/in/royashbrook
|blog|https://ashbrook.io

# v7.3.0

- Fix throwing error when a Worksheet name collides with a method, or property name on the `OfficeOpenXml.ExcelPackage` package

# v7.2.3

- Fix inline help, thank you [Wes Stahler](https://github.com/stahler)

# v7.2.2

- Improved checks for Linux, Mac and PS 5.1

# v7.2.1

- Improve auto-detection of data on the clipboard

# v7.2.0

- Added `Read-Clipboard` support for Windows. Read text from clipboard. It can read CSV or JSON. Plus, you can specify the delimiter and headers.

### Check out the video
<a href="https://youtu.be/dv2GOH5sbpA"><img src="https://img.youtube.com/vi/dv2GOH5sbpA/0.jpg" width="400"></a>