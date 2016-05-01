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

    Context "Only length in the Excel sheet, not the string #41" {
        # https://github.com/dfinke/ImportExcel/issues/41

        $xlPkg = "test", "c:\whatever", "d:\whatever" | Export-Excel -Path $workbook -DateTimeFormat $cultureShortDatePattern -PassThru
        It "Can accept an array of strings" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $cell = $ws.Cells["A1"]
            $cell.Value | Should Be "test" 
            $cell = $ws.Cells["A2"]
            $cell.Value | Should Be "c:\whatever" 
            $cell = $ws.Cells["A3"]
            $cell.Value | Should Be "d:\whatever" 
        }
        $xlPkg.Save(); $xlPkg.Dispose();
        # Invoke-Item $workbook; throw;

    }

    Context "Numbers with leading zeros treated as numbers, not text #33" {
        # https://github.com/dfinke/ImportExcel/issues/33

        $csvData = 
@"
a,b,c
01,002,3
00u812,05150,abc
123,456,098
0123,456,098
"@ | ConvertFrom-Csv

        $xlPkg = $csvData | Export-Excel -Path $workbook -PassThru
        It "Can accept CSV data with numeric strings that have leading zeroes" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $cell = $ws.Cells["A2"]
            $cell.Value -is [string] | Should Be $true 
            $cell.Value | Should Be "01" 
            $cell = $ws.Cells["A3"]
            $cell.Value -is [string] | Should Be $true 
            $cell.Value | Should Be "00u812" 
            $cell = $ws.Cells["A4"]
            $cell.Value -is [double] | Should Be $true 
            $cell.Value | Should Be 123 
            $cell = $ws.Cells["A5"]
            $cell.Value | Should Be "0123" 
            $cell = $ws.Cells["B2"]
            $cell.Value | Should Be "002" 
            $cell = $ws.Cells["B3"]
            $cell.Value | Should Be "05150" 
            $cell = $ws.Cells["B4"]
            $cell.Value -is [double] | Should Be $true 
            $cell.Value | Should Be 456 
            $cell = $ws.Cells["B5"]
            $cell.Value -is [double] | Should Be $true 
            $cell.Value | Should Be 456 
            $cell = $ws.Cells["C4"]
            $cell.Value | Should Be "098" 
            $cell = $ws.Cells["C5"]
            $cell.Value | Should Be "098" 
        }
        $xlPkg.Save(); $xlPkg.Dispose();
        # Invoke-Item $workbook; throw;

        $columnOptions = @{ "*" = @{ IgnoreText = $true } }
        $xlPkg = $csvData | Export-Excel -Path $workbook -ColumnOptions $columnOptions -PassThru
        It "Can skip automatic conversion of CSV strings using -ColumnOptions and IgnoreText" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $cols = $ws.Cells["A2:C"] # Skip the headings.
            $cols | % {
                $cell = $_
                if ($cell -ne $null) {
                    $cell.Value -is [string] | Should Be $true
                }
            }

            $cell = $ws.Cells["A2"]
            $cell.Value -is [string] | Should Be $true 
            $cell.Value | Should Be "01" 
            $cell = $ws.Cells["A3"]
            $cell.Value -is [string] | Should Be $true 
            $cell.Value | Should Be "00u812" 
            $cell = $ws.Cells["A4"]
            $cell.Value -is [string] | Should Be $true 
            $cell.Value | Should Be "123"
            $cell = $ws.Cells["A5"]
            $cell.Value | Should Be "0123" 
            $cell = $ws.Cells["B2"]
            $cell.Value | Should Be "002" 
            $cell = $ws.Cells["B3"]
            $cell.Value | Should Be "05150" 
            $cell = $ws.Cells["B4"]
            $cell.Value -is [string] | Should Be $true 
            $cell.Value | Should Be "456" 
            $cell = $ws.Cells["B5"]
            $cell.Value -is [string] | Should Be $true 
            $cell.Value | Should Be "456"
            $cell = $ws.Cells["C4"]
            $cell.Value | Should Be "098" 
            $cell = $ws.Cells["C5"]
            $cell.Value | Should Be "098" 
        }
        $xlPkg.Save(); $xlPkg.Dispose();
        # Invoke-Item $workbook; throw;
    }

    Remove-TestWorkbook
}
