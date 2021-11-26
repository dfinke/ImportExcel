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
Contents of file.xlsx after modifying:  
![ExcelFileContentsPostAdd](/images/FAQ_Images/ExcelFileContentsPostAdd.png)
Can also load a value at a specific index:
```powershell
$ValueAtIndex = $WorkSheet[2,1].Value #Loads the value at row 2, column A
```
```$ValueAtIndex``` now contains: ![ValueAtIndexData](/images/FAQ_Images/ValueAtIndexData.png)  
### Save File After Modifying
The changes will not display in the Excel file until Close-ExcelPackage is called.  
```powershell
Close-ExcelPackage $ExcelPkgFile #close and save changes made to the Excel file.
```
**Note**: If the file is currently in use, Close-ExcelPackage will return an error and will not save the information.

