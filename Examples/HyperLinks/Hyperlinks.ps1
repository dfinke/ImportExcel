try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

@"
site,link
google,http://www.google.com
stackoverflow,http://stackoverflow.com
microsoft,http://microsoft.com
"@ | ConvertFrom-Csv | Export-Excel
