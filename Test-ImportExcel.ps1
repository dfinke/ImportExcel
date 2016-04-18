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

PS> Test-ImportExcel

Runs every Pester test in the module (i.e. all files named *.Tests.ps1).

.EXAMPLE

PS> Test-ImportExcel -TestName "Export-Excel"

Runs the Pester test in the module that describes itself as "Export-Excel".

#>
function Test-ImportExcel {
    [CmdletBinding()]
    param([string]$TestName)
    Push-Location $PSScriptRoot
    try {
        Invoke-Pester -TestName $TestName
    }
    catch {
    }
    finally {
        Pop-Location
    }
}
