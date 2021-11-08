#Requires -Modules Pester
if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}

$skip = $false
if ($IsLinux -or $IsMacOS) {
    $skip = $true
    Write-Warning "Invoke-ExcelQuery: Linux and MacOs are not supported. Skipping tests."
}else{
    try {
        if ((New-Object system.data.oledb.oledbenumerator).GetElements().SOURCES_NAME -notcontains "Microsoft.ACE.OLEDB.12.0") {
            $skip = $true
            Write-Warning "Invoke-ExcelQuery: Microsoft.ACE.OLEDB.12.0 provider not found. Skipping tests."
        }
    }
    catch {
        $skip = $true
        Write-Warning "Invoke-ExcelQuery: Calls to System.Data.OleDb failed. Skipping tests."
    }
}


Describe "Invoke-ExcelQuery" -Tag "Invoke-ExcelQuery" {
    $PSDefaultParameterValues = @{ 'It:Skip' = $skip }
    BeforeAll {
        $tfp = "$PSScriptRoot\Read-OleDbData.xlsx"
    }
    Context "Basic Checks" {
        It "Should have a valid Test file" {
            Test-Path $tfp | Should -Be $true
        }
        It "Should have the Read-OleDbData command loaded" {
            (Get-Command Read-OleDbData -ErrorAction SilentlyContinue) -ne $null | Should -Be $true
        }
        It "Should have the Invoke-ExcelQuery command loaded" {
            (Get-Command Invoke-ExcelQuery -ErrorAction SilentlyContinue) -ne $null | Should -Be $true
        }
    }
    Context "Sheet1`$A1" {
        It "Should return 1 result with a value of 1" {
            $Results = Invoke-ExcelQuery $tfp "select ROUND(F1) as [A1] from [sheet1`$A1:A1]"
            @($Results).length + $Results.A1 | Should -Be 2
        }
    }
}
