#Requires -Module Pester
Set-StrictMode -Version Latest

. $PSScriptRoot\TestHelpers.ps1

Describe "ImportExcel" {
    Context "Performing sanity check" {
        It "Verifying test helpers" {
            Get-Command New-TestWorkbook | Should Not Be $null
            Get-Command Remove-TestWorkbook | Should Not Be $null
            Get-Command New-TestDataCsv | Should Not Be $null
            Get-Command Get-PathEPPlusDll | Should Not Be $null
            Get-Command Add-TypeEPPlus | Should Not Be $null
        }
        It "Verifying EPPlus.dll" {
            Get-PathEPPlusDll | Should Not Be $null
        }
    }
    Context "Importing CSV data from a here string" {
        $csvData = New-TestDataCsv
        It "All properties are type [string]" {
            $csvData | % {
                $_.PSObject.Properties | % {
                    $_.Value -is [string] | Should Be $true
                }
            }
        }
        It "Leading zeroes are preserved" {
            $csvData[4] | Select-Object -ExpandProperty ID | Should Be "00120"
        }
    }
}
