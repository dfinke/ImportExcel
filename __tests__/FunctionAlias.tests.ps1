#Requires -Modules Pester
remove-module importExcel -erroraction silentlyContinue
Import-Module $PSScriptRoot\..\ImportExcel.psd1 -Force


Describe "Check if Function aliases exist" {

    It "Set-Column should exist" {
        ${Alias:Set-Column} | Should Not BeNullOrEmpty
    }

    It "Set-Row should exist" {
          ${Alias:Set-Row} | Should Not BeNullOrEmpty
    }

    It "Set-Format should exist" {
          ${Alias:Set-Format} | Should Not BeNullOrEmpty
    }

  <#It "Merge-MulipleSheets should exist" {
        Get-Command Merge-MulipleSheets | Should Not Be $null
    }
#>
    It "New-ExcelChart should exist" {
          ${Alias:New-ExcelChart} | Should Not BeNullOrEmpty
    }

}