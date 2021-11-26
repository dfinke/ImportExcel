# Write to an Existing Excel File
### Enumerate the Excel File
```powershell
$ExcelPkgFile = Open-ExcelPackage -Path  "C:\Test\file.xlsx"
```
Contents of file.xlsx:  
![ExcelFileContents](/images/FAQ_Images/ExcelFileContents.png)
### Enumerate the Worksheet to View or Modify the Data
```powershell
$WorkSheet = $ExcelPkgFile.Workbook.Worksheets["sheet1"].Cells #open excel worksheet cells from worksheet "sheet1"
```
Visual of Data Structure:  
![DataStructureExcelPkg](/images/FAQ_Images/DataStructureExcelPkg.png)

Modify a specific value by accessing row/col like a 2D Array:
```powershell
$WorkSheet[1,4].Value = "New Column Header" #Starts at index 1 not 0
```

