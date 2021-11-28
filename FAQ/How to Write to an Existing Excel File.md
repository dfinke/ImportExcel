# Write to an Existing Excel File
### Enumerate the Excel File
The cmdlets ```Open-ExcelPackage``` and ```Close-ExcelPackage``` allow for direct modification to Excel file contents.
```powershell
$ExcelPkg = Open-ExcelPackage -Path  "C:\Test\file.xlsx"
```
Contents of file.xlsx:  
![ExcelFileContents](/images/FAQ_Images/ExcelFileContents.png)
### Enumerate the Worksheet to View or Modify the Data
```powershell
$WorkSheet = $ExcelPkg.Workbook.Worksheets["sheet1"].Cells #open excel worksheet cells from worksheet "sheet1"
```
Visual of data structure:  
![DataStructureExcelPkg](/images/FAQ_Images/DataStructureExcelPkg.png)  
A1 contains "someHeader", A2 contains "data1" etc.
### Modify a Specific Value in a File 
Values can be accessed by row, column. Similar to a 2D array.
```powershell
$WorkSheet[1,4].Value = "New Column Header" #Starts at index 1 not 0
```
Contents of file.xlsx after modifying:  
![ExcelFileContentsPostAdd](/images/FAQ_Images/ExcelFileContentsPostAdd.png)  
### Load Value at Specific Index
```powershell
$ValueAtIndex = $WorkSheet[2,1].Value #Loads the value at row 2, column A
```
```$ValueAtIndex``` now contains: ![ValueAtIndexData](/images/FAQ_Images/ValueAtIndexData.png)  
### Save File After Modifying
The changes will not display in the Excel file until Close-ExcelPackage is called.  
```powershell
Close-ExcelPackage $ExcelPkg #close and save changes made to the Excel file.
```
**Note**: If the file is currently in use, Close-ExcelPackage will return an error and will not save the information.

