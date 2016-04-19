<#
.SYNOPSIS

Runs the Pester tests for the ImportExcel module.

.DESCRIPTION

This function depends on the Pester module, that can be installed via
Install-Module. It may be necessary to do an Import-Module Pester
before running this function. Alternatively, add Pester to the
RequiredModules section of your own script module (.psd1).

If changes are made to the ImportExcel module, it may be necessary to
run Import-Module ImportExcel -Force, before running this function, to
ensure that the latest changes are tested.

.EXAMPLE

PS> & .\TestImportExcel.ps1

Runs every Pester test in the module (i.e. all files named *.Tests.ps1).

.EXAMPLE

PS> & .\TestImportExcel.ps1 -TestName "Export-Excel"

Runs the Pester test in the module that describes itself as "Export-Excel".

#>
[CmdletBinding()]
param([string]$TestName)
Set-StrictMode -Version Latest
Push-Location $PSScriptRoot
$invokePester = Get-Command Invoke-Pester -ErrorAction SilentlyContinue
if ($invokePester -eq $null) {
    $installModule = Get-Command Install-Module -ErrorAction SilentlyContinue
    if ($installModule -ne $null) {
        "Pester is not installed. Press any key to run 'Install-Module Pester'." | Out-Host
        $Host.UI.RawUI.ReadKey() | Out-Null
        "Installing Pester..." | Out-Host
        Install-Module Pester -Scope CurrentUser -Verbose
    }
    else {
        throw "Pester is not installed."
    }
}
try {
    "Running tests..." | Out-Host
    Invoke-Pester -TestName $TestName
}
catch {
    throw $_
}
finally {
    Pop-Location
}
