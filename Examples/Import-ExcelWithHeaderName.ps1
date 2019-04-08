# Create Test File
$path = "$env:TEMP\Test.xlsx"
Remove-item -Path $path  -ErrorAction SilentlyContinue
$processes = Get-Process | Select-Object -First 10
$Processes | Export-Excel $path

# Import the file as is.
Import-Excel -Path $path | Format-Table
# Import the file replacing the name of the 2 first headers and discarding the rest.
Import-Excel -Path $path -HeaderName 'NewA', 'NewB'
# Import the file selecting 3 of it's columns by name, renaming 2, changing the order and discarding the rest.
Import-Excel -Path $path -HeaderName ([Ordered]@{Name = 'Process Name' ; StartTime = 'Start'; Path = 'Path'})
<# Output Example:
Process Name         Start               Path
------------         -----               ----
ApplicationFrameHost 07/04/2019 11:02:19 C:\WINDOWS\system32\ApplicationFrameHost.exe
AppVShNotify
audiodg              08/04/2019 19:21:32
chrome               08/04/2019 16:49:47 C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
chrome               08/04/2019 16:49:27 C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
chrome               08/04/2019 16:49:27 C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
#>