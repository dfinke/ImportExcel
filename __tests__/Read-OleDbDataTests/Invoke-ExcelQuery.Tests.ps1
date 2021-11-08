#Requires -Modules Pester
$scriptPath = $PSScriptRoot
Import-Module $scriptPath\..\..\ImportExcel.psd1 -Force
$tfp = "$scriptPath\Read-OleDbData.xlsx"
$skip = $IsLinux -or $IsMacOS #init default, not supported on mac or linux
try {
    $IsMissingACE = $null -eq ((New-Object system.data.oledb.oledbenumerator).GetElements().SOURCES_NAME -like "Microsoft.ACE.OLEDB*")    
    if ($IsMissingACE) {
        Write-Warning "MICROSOFT.ACE.OLEDB is missing! Tests will be skipped."
    }
    $skip = $skip -and $IsMissingACE
}
catch {
    Write-Warning "Unable to get sources from System.Data.OleDb. Tests will be skipped."
    $skip = $true #this will fail if the call to get the sources fails, usually means System.Data.OleDb isn't installed/supported
}


Write-Warning "`$tfp = '$tfp'"
Write-Warning "`Test-Path $tfp = '$(Test-Path $tfp)'"
Write-Warning "`$IsMissingACE = '$IsMissingACE'"

Describe "Invoke-ExcelQuery" -Tag "Invoke-ExcelQuery" {
    $PSDefaultParameterValues = @{ 'It:Skip' = $skip }
    Context "Basic Checks" {
        It "Should have a valid Test file" {
            Test-Path $tfp | Should -Be $true
        }
        It "Should have the Read-OleDbData command loaded" {
            (Get-Command Read-OleDbData) -ne $null | Should -Be $true
        }
        It "Should have the Invoke-ExcelQuery command loaded" {
            (Get-Command Invoke-ExcelQuery) -ne $null | Should -Be $true
        }
    }
    Context "Sheet1`$A1" {
        It "Should return 1 result with a value of 1" {
            $Results = Invoke-ExcelQuery $tfp "select ROUND(F1) as [A1] from [sheet1`$A1:A1]"
            @($Results).length + $Results.A1 | Should -Be 2
        }
    }
}
