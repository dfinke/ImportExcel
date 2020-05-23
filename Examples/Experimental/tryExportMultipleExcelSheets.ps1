. "$PSScriptRoot\Export-MultipleExcelSheets.ps1"

$xlfile = "$env:TEMP\test.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$p = Get-Process

$InfoMap = @{
    PM                 = { $p | Select-Object company, pm }
    Handles            = { $p | Select-Object company, handles }
    Services           = { Get-Service }
    Files              = { Get-ChildItem -File }
    Albums             = { ConvertFrom-Csv (Invoke-RestMethod https://raw.githubusercontent.com/dfinke/powershell-for-developers/master/chapter05/BeaverMusic/BeaverMusic.UI.Shell/albums.csv) }
    WillNotGetExported = "Hello World"
}

Export-MultipleExcelSheets -Path $xlfile -InfoMap $InfoMap -Show -AutoSize