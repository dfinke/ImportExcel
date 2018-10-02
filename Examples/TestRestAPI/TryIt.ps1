try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

. $PSScriptRoot\TestAPIReadXls.ps1

Test-APIReadXls $PSScriptRoot\testlist.xlsx | Foreach-Object {
    Invoke-Pester -Script $_.fullname  -PassThru -Show None
}