# How to Read an existing Excel File

```powershell
Import-Module ImportExcel
#Loads the Excel file into a is a custom PS Object
$ExcelFile = Import-Excel "C:\Test\file.xlsx" -WorksheetName "Sheet1" 
```

## Visual of Data Structure
The File C:\Test\file.xlsx contains  
![alt text](https://github.com/DavisHenckel/ImportExcel/blob/FAQ_Docs/images/FAQ_Images/ExcelFileContents.png)

After Loading this data into ```$ExcelFile``` the data is stored like:  
![alt text](https://github.com/DavisHenckel/ImportExcel/blob/FAQ_Docs/images/FAQ_Images/ExcelFileDebugImg.jpg)
