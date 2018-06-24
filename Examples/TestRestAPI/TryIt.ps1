. $PSScriptRoot\TestAPIReadXls.ps1

Test-APIReadXls $PSScriptRoot\testlist.xlsx | % {
    Invoke-Pester -Script $_.fullname  -PassThru -Show None
}