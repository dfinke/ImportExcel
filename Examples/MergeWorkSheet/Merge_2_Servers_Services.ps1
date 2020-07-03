try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

Remove-Item -Path  "$env:temp\server*.xlsx" , "$env:temp\Combined*.xlsx" -ErrorAction SilentlyContinue

#Get a subset of services into $s and export them
[System.Collections.ArrayList]$s = Get-service | Select-Object -first 25 -Property *
$s | Export-Excel -Path $env:temp\server1.xlsx

#$s is a zero based array, excel rows are 1 based and excel has a header row so Excel rows will be 2 + index in $s.
#Change a row. Add a row. Delete a row.  And export the changed $s to a second file.
$s[2].DisplayName = "Changed from the orginal"   #This will be row 4 in Excel - this should be highlighted as a change

$d = $s[-1] | Select-Object -Property *
$d.DisplayName = "Dummy Service"
$d.Name = "Dummy"
$s.Insert(3,$d)                                 #This will be row 5 in Excel - this should be highlighted as a new item

$s.RemoveAt(5)                                  #This will be row 7 in Excel - this should be highlighted as deleted item

$s | Export-Excel -Path $env:temp\server2.xlsx

#This use of Merge-worksheet Assumes a default worksheet name, (sheet1)  We will check and output Name (the key), DisplayName and StartType and ignore other properties.
Merge-Worksheet -Referencefile "$env:temp\server1.xlsx" -Differencefile  "$env:temp\Server2.xlsx" -OutputFile  "$env:temp\combined1.xlsx"  -Property name,displayname,startType -Key name -Show