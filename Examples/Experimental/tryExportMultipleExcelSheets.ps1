. "$PSScriptRoot\Export-MultipleExcelSheets.ps1"

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

$p = Get-Process

$InfoMap = @{
    PM                 = { $p | Select-Object company, pm }
    Handles            = { $p | Select-Object company, handles }
    Services           = { Get-Service }
    Files              = { Get-ChildItem -File }
    Albums             = { ConvertFrom-Csv (Invoke-RestMethod https://raw.githubusercontent.com/dfinke/powershell-for-developers/master/chapter05/BeaverMusic/BeaverMusic.UI.Shell/albums.csv) }
    WillNotGetExported = "Hello World"
}

Export-MultipleExcelSheets -Path $xlSourcefile -InfoMap $InfoMap -Show -AutoSize