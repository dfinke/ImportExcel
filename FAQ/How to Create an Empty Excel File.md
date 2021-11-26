# Create an Empty Excel File
Use an empty string to export to an excel File.
```powershell
#Build an Excel file named: "file.xlsx" containing a worksheet: "MyWorksheet"
"" | Export-Excel -Path "C:\Test\file.xlsx -WorksheetName "MyWorksheet" 
```
