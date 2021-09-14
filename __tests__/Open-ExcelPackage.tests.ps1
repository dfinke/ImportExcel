#Requires -Modules Pester

if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}

<#
Methods
-------
Dispose
Equals
GetAsByteArray
GetHashCode
GetType
Load
Save
SaveAs
ToString

Properties
----------

Compatibility
Compression
DoAdjustDrawings
Encryption
File
Package
Stream
Workbook
#>

Describe "Test Open Excel Package" -Tag Open-ExcelPackage { 
    It "Should handle opening a workbook with Worksheet Names that will cause errors" {
        $xlFilename = "$PSScriptRoot\UnsupportedWorkSheetNames.xlsx"

        { Open-ExcelPackage -Path $xlFilename -ErrorAction Stop } | Should -Not -Throw
    }
}