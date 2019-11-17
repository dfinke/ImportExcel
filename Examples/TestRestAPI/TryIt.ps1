try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

. $PSScriptRoot\TestAPIReadXls.ps1

Test-APIReadXls $PSScriptRoot\testlist.xlsx | Foreach-Object {
    Invoke-Pester -Script $_.fullname  -PassThru -Show None
}