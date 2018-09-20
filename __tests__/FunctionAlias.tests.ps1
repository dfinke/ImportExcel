#Requires -Modules Pester
Import-Module $PSScriptRoot\..\ImportExcel.psd1 -Force

Describe "Check if Function aliases exist" {

    It "Set-Column should exist" {
        Get-Command Set-Column | Should Not Be $null
    }

    It "Set-Row should exist" {
        Get-Command Set-Row | Should Not Be $null
    }

    It "Set-Format should exist" {
        Get-Command Set-Format | Should Not Be $null
    }

    It "Merge-MulipleSheets should exist" {
        Get-Command Merge-MulipleSheets | Should Not Be $null
    }

    It "New-ExcelChart should exist" {
        Get-Command New-ExcelChart | Should Not Be $null
    }

}