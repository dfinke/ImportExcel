Import-Module $PSScriptRoot -Force -Scope Global

function New-TestWorkbook {
    $testWorkbook = "$($PSScriptRoot)\issues.xlsx"
    
    Remove-Item $testWorkbook -ErrorAction Ignore
    $testWorkbook
}

function Remove-TestWorkbook {
    New-TestWorkbook | Out-Null
}

Describe "Issues" {

    $workbook = New-TestWorkbook

    Context "Keep numbers as strings formatted as text #92" {
        # https://github.com/dfinke/ImportExcel/issues/92

        $xlPkg = [pscustomobject]@{PhoneNumber=[string]"01234123456"} | Export-Excel -Path $workbook -PassThru
        It "Keeps numeric strings with leading zeroes as text" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $cell = $ws.Cells["A2"]
            $cell.Value -is [string] | Should Be $true
            $cell.Value | Should Be "01234123456"
        }
        $xlPkg.Save(); $xlPkg.Dispose();
        # Invoke-Item $workbook; throw;

        $columnOptions = @{ PhoneNumber = @{ IgnoreText = $true } }
        $xlPkg = [pscustomobject]@{PhoneNumber=[string]"1234123456"} | Export-Excel -Path $workbook -ColumnOptions $columnOptions -PassThru
        It "Can ignore the automatic conversion of strings using -ColumnOptions and IgnoreText" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $cell = $ws.Cells["A2"]
            $cell.Value -is [string] | Should Be $true
            $cell.Value | Should Be "1234123456"
        }
        $xlPkg.Save(); $xlPkg.Dispose();
        # Invoke-Item $workbook; throw;

        $columnOptions = @{ PhoneNumber = @{ ForceText = $true } }
        $xlPkg = [pscustomobject]@{PhoneNumber=1234123456} | Export-Excel -Path $workbook -ColumnOptions $columnOptions -PassThru
        It "Can force numbers as strings using -ColumnOptions and ForceText" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $cell = $ws.Cells["A2"]
            $cell.Value -is [string] | Should Be $true
            $cell.Value | Should Be "1234123456"
        }
        $xlPkg.Save(); $xlPkg.Dispose();
        # Invoke-Item $workbook; throw;
    }

    Context "Use localized date format in Export-Excel #52" {
        # https://github.com/dfinke/ImportExcel/issues/52

        $cultureShortDatePattern = [CultureInfo]::CurrentCulture.DateTimeFormat.ShortDatePattern

        $xlPkg = "$(Get-Date)" | Export-Excel -Path $workbook -DateTimeFormat $cultureShortDatePattern -PassThru
        It "Can accept localized date format with -DateTimeFormat" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $cell = $ws.Cells["A1"]
            $cell.Value -is [DateTime] | Should Be $true
            $cell.Style.NumberFormat.Format | Should Be $cultureShortDatePattern
        }
        $xlPkg.Save(); $xlPkg.Dispose();
        # Invoke-Item $workbook; throw;

        $columnOptions = @{ 1 = @{ DateTimeFormat = $cultureShortDatePattern } }
        $xlPkg = "$(Get-Date)" | Export-Excel -Path $workbook -ColumnOptions $columnOptions -PassThru
        It "Can accept localized date format with -ColumnOptions and DateTimeFormat" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $cell = $ws.Cells["A1"]
            $cell.Value -is [DateTime] | Should Be $true
            $cell.Style.NumberFormat.Format | Should Be $cultureShortDatePattern
        }
        $xlPkg.Save(); $xlPkg.Dispose();
        # Invoke-Item $workbook; throw;
    }

    Remove-TestWorkbook
}
