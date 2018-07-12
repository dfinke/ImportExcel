try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

@"
site,link
google,http://www.google.com
stackoverflow,http://stackoverflow.com
microsoft,http://microsoft.com
"@ | ConvertFrom-Csv | Export-Excel
