try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

Remove-Item -Path  "$env:temp\server*.xlsx" , "$env:temp\Combined*.xlsx" -ErrorAction SilentlyContinue

#Get a subset of services into $s and export them
[System.Collections.ArrayList]$s = Get-service | Select-Object -first 25 -Property Name,DisplayName,StartType
$s | Export-Excel -Path $env:temp\server1.xlsx

#$s is a zero based array, excel rows are 1 based and excel has a header row so Excel rows will be 2 + index in $s.
#Change a row. Add a row. Delete a row.  And export the changed $s to a second file.
$row4Displayname  = $s[2].DisplayName
$s[2].DisplayName = "Changed from the orginal"    #This will be excel row 4 and Server 2 will show as changed.

$d = $s[-1] | Select-Object -Property *
$d.DisplayName = "Dummy Service"
$d.Name = "Dummy"
$s.Insert(3,$d)                                   #This will be Excel row 5 and server 2 will show as changed - so will Server 3

$s.RemoveAt(5)                                    #This will be Excel row 7 and Server 2 will show as missing.

$s | Export-Excel -Path $env:temp\server2.xlsx

#Make some more changes to $s and export it to a third file
$s[2].displayname = $row4Displayname             #Server 3 row 4 will match server 1 so won't be highlighted

$d = $s[-1] | Select-Object -Property *
$d.DisplayName = "Second Service"
$d.Name = "Service2"
$s.Insert(6,$d)                                  #This will be an extra row in Server 3 at row 8. It will show as missing in Server 2.
$s.RemoveAt(8)                                   #This will show as missing in Server 3 at row 11 ()

$s | Export-Excel -Path $env:temp\server3.xlsx

#Now bring the three files together.

Merge-MultipleSheets -Path "$env:temp\server1.xlsx", "$env:temp\Server2.xlsx","$env:temp\Server3.xlsx" -OutputFile "$env:temp\combined3.xlsx"  -Property name,displayname,startType -Key name  -Show