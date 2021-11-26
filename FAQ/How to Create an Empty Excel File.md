# Create an Empty Excel File
Use an Empty String and Export to an Excel File.
```powershell
#Build an Excel file named: "file.xlsx" containing a worksheet: "MyWorksheet"
"" | Export-Excel -Path "C:\Test\file.xlsx -WorksheetName "MyWorksheet" 
```
