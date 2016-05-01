#Requires -Module Pester
#Requires -Module ImportExcel
Set-StrictMode -Version Latest

Import-Module $PSScriptRoot -Force -Scope Global

# Bring New-CellData helpers into scope.
. (Join-Path $PSScriptRoot "$(Split-Path -Leaf $PSCommandPath)".Replace(".Tests.ps1", ".ps1"))

function New-TestWorkbook {
    $testWorkbook = Join-Path $PSScriptRoot test.xlsx
    if (Test-Path $testWorkbook) {
        rm $testWorkbook -Force
    }
    $testWorkbook
}

function Remove-TestWorkbook {
    New-TestWorkbook | Out-Null
}

function Get-DateFormatDefault {
    "mmm/dd/yyyy hh:mm:ss"
}

Describe "DoubleTryParse" {
    Context "Parsing decimal strings" {

        $commaCulture = (0.3 | Out-String -Stream) -eq "0,3"

        It "Converts 0.1 to 0.1 with Any/InvariantInfo" {
            $double = 0
            # https://msdn.microsoft.com/en-us/library/system.globalization.numberstyles(v=vs.110).aspx
            [double]::TryParse("0.1", [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::InvariantInfo, [ref]$double) | Should Be $true
            $double | Should Be 0.1
            "$double" | Should Be "0.1"
            if ($commaCulture -eq $true) {
                $double | Out-String -Stream | Should Be "0,1"
            }
        }
        It "Converts 0,1 to 1 with Any/InvariantInfo" {
            $double = 0
            [double]::TryParse("0,1", [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::InvariantInfo, [ref]$double) | Should Be $true
            $double | Should Be 1
            "$double" | Should Be "1"
            $double | Out-String -Stream | Should Be "1"
        }
        It "Converts 0,3 to 3 with Any/InvariantInfo" {
            $double = 0
            [double]::TryParse("0,3", [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::InvariantInfo, [ref]$double) | Should Be $true
            $double | Should Be 3
            "$double" | Should Be "3"
            if ($commaCulture -eq $true) {
                $double | Out-String -Stream | Should Be "3"
            }
        }

        if ($commaCulture -eq $true) {
            It "Converts 0,3 to 0,3 with Any/CurrentInfo" {
                $double = 0
                [double]::TryParse("0,3", [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::CurrentInfo, [ref]$double) | Should Be $true
                $double | Should Be 0.3
                "$double" | Should Be "0.3"
                $double | Out-String -Stream | Should Be "0,3"
            }
            It "Fails to convert 0.3 to 0,3 with Any/CurrentInfo" {
                $double = 0
                [double]::TryParse("0.3", [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::CurrentInfo, [ref]$double) | Should Be $false
                $double | Should Be 0
                "$double" | Should Be "0"
            }
        }
    }
}

Describe "NewCellData" {

    Context "Piping [string] inputs" {

        It "Converts numeric strings to [double]" {
            New-CellData -TargetData "12345" | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 12345
                $_.Format | Should Be "General"
            }
        }

        It "Leaves numeric strings with leading zeroes as strings" {
            New-CellData -TargetData "012345" | % {
                $_.Value -is [string] | Should Be $true
                $_.Value | Should Be "012345"
                $_.Format | Should Be "General"
            }
        }

        It "Numeric strings with leading zeroes that are non-integers are treated as numbers" {
            New-CellData -TargetData "0.01" | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 0.01
                $_.Format | Should Be "General"
            }
        }

        It "Leaves numeric strings as text when using -ForceText switch" {
            New-CellData -TargetData "12345" -ForceText | % {
                $_.Value -is [string] | Should Be $true
                $_.Value | Should Be "12345"
                $_.Format | Should Be "General"
            }
        }

        It "Converts date strings to the default date format" {
            $date = Get-Date
            New-CellData -TargetData "$date" | % {
                $_.Value -is [DateTime] | Should Be $true
                "$($_.Value)" | Should Be "$date"
                $_.Format | Should Be (Get-DateFormatDefault)
            }
        }

        It "Leaves numeric strings with starting and trailing whitespace as strings" {
            New-CellData -TargetData " 12345" | % {
                $_.Value -is [string] | Should Be $true
                $_.Value | Should Be " 12345"
                $_.Format | Should Be "General"
            }
            New-CellData -TargetData "12345 " | % {
                $_.Value -is [string] | Should Be $true
                $_.Value | Should Be "12345 "
                $_.Format | Should Be "General"
            }
        }

        It "Keeps percentage strings as text" {
            New-CellData -TargetData "90%" | % {
                $_.Value -is [string] | Should Be $true
                $_.Value | Should Be "90%"
                $_.Format | Should Be "General"
            }
        }

        $nfiZa = [CultureInfo]::GetCultureInfo("en-ZA").NumberFormat

        It "Converts 12,34 to 12.34 with -NumberFormatInfo ZA" {
            New-CellData -TargetData "12,34" -NumberFormatInfo $nfiZa | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 12.34
                $_.Format | Should Be "General"
            }
        }

        It "Converts 12.34 to 12.34" {
            New-CellData -TargetData "12.34" | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 12.34
                $_.Format | Should Be "General"
            }
        }

        It "Converts 12,345 to 12.345 with -NumberFormatInfo ZA" {
            New-CellData -TargetData "12,345" -NumberFormatInfo $nfiZa | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 12.345
                $_.Format | Should Be "General"
            }
        }

        It "Converts 12.345 to 12.345" {
            New-CellData -TargetData "12.345" | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 12.345
                $_.Format | Should Be "General"
            }
        }

        It "Converts 0,1 to 0.1 with -NumberFormatInfo ZA" {
            New-CellData -TargetData "0,1" -NumberFormatInfo $nfiZa | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 0.1
                $_.Format | Should Be "General"
            }
        }

        It "Converts 0.1 to 0.1" {
            New-CellData -TargetData "0.1" | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 0.1
                $_.Format | Should Be "General"
            }
        }

        It "Converts 0,01 to 0.01 with -NumberFormatInfo ZA" {
            New-CellData -TargetData "0,01" -NumberFormatInfo $nfiZa | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 0.01
                $_.Format | Should Be "General"
            }
        }

        It "Converts 0.01 to 0.01" {
            New-CellData -TargetData "0.01" | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 0.01
                $_.Format | Should Be "General"
            }
        }

    }

    Context "Piping [DateTime] inputs" {
        It "Outputs the same [DateTime]" {
            $date = Get-Date
            New-CellData -TargetData $date | % {
                $_.Value -is [DateTime] | Should Be $true
                $_.Value | Should Be $date
                $_.Format | Should Be (Get-DateFormatDefault)
            }
        }
    }

    Context "Piping numeric value type inputs" {
        It "Outputs [int] for [int] input" {
            New-CellData -TargetData 123 | % {
                $_.Value -is [int] | Should Be $true
                $_.Value | Should Be 123
                $_.Format | Should Be "General"
            }
        }
        It "Outputs [double] for [double] input" {
            New-CellData -TargetData ([double]123) | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 123
                $_.Format | Should Be "General"
            }
        }
        It "Outputs [long] for [long] input" {
             New-CellData -TargetData ([long]123) | % {
                $_.Value -is [long] | Should Be $true
                $_.Value | Should Be 123
                $_.Format | Should Be "General"
            }
        }
    }

    Context "Piping other value type inputs" {
        It "Outputs [bool] for [bool] input" {
            New-CellData -TargetData $true | % {
                $_.Value -is [bool] | Should Be $true
                $_.Value | Should Be $true
                "$($_.Value)" | Should Be "True"
                $_.Format | Should Be "General"
            }
        }
    }

    Context "Piping CSV data" {
        $csvData = @"
        Name, ID, Age, Birthday
        Aa, 123, 82, 12 January 1984
        BB, 012, 34, 12 August 1955
        CC, 901, 44, 30 May 1801
"@ | ConvertFrom-Csv
        It "Converts property values to appropriate types" {
            $csvData | Select-TargetData Name | New-CellData | % {
                $_.Value -is [string] | Should Be $true
                $_.Format | Should Be "General"
            }
            $csvData | Select-TargetData ID | New-CellData -ForceText | % {
                $_.Value -is [string] | Should Be $true
                $_.Format | Should Be "General"
            }
            $csvData | Select-TargetData Age | New-CellData | % {
                $_.Value -is [double] | Should Be $true
                $_.Format | Should Be "General"
            }
            $csvData | Select-TargetData Age | New-CellData -IgnoreText | % {
                $_.Value -is [string] | Should Be $true
                $_.Format | Should Be "General"
            }
            $csvData | Select-TargetData Birthday | New-CellData | % {
                $_.Value -is [DateTime] | Should Be $true
                $_.Format | Should Be (Get-DateFormatDefault)
            }
            $csvData | Select-TargetData Birthday | New-CellData -ForceText | % {
                $_.Value -is [string] | Should Be $true
                $_.Format | Should Be "General"
            }
        }
    }

    Context "Piping Get-Process data" {
        $process = Get-Process powershell
        It "Converts property values to appropriate types" {
            $process | Select-TargetData StartTime | New-CellData | % {
                $_.Value -is [DateTime] | Should Be $true
                $_.Format | Should Be (Get-DateFormatDefault)
            }
            $process | Select-TargetData Id | New-CellData | % {
                $_.Value -is [int] | Should Be $true
                $_.Format | Should Be "General"
            }
            $process | Select-TargetData ProcessName | New-CellData | % {
                $_.Value -is [string] | Should Be $true
                $_.Format | Should Be "General"
            }
            $process | Select-TargetData Handles | New-CellData | % {
                $_.Value -is [int] | Should Be $true
                $_.Format | Should Be "General"
            }
        }
    }

    Context "With Export-Excel and PsCustomObject" {
        $workbook = New-TestWorkbook
        $process = Get-Process powershell | Select-Object Id, StartTime, PriorityClass, TotalProcessorTime
        $xlPkg = $process | Export-Excel $workbook -PassThru
        It "Produces correctly formatted sheet for Get-Process" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $col = $ws.Cells["A2:A"] # Id
            $col | % {
                $_.Value -is [int] | Should Be $true
                $_.Style.NumberFormat.Format | Should Be "General"
            }
            $col = $ws.Cells["B2:B"] # StartTime
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [DateTime] | Should Be $true
            }
            $col | % {
                if ($_.Value -ne $null) {
                    $_.Style.NumberFormat.Format | Should Be (Get-DateFormatDefault)
                }
                else {
                    $_.Style.NumberFormat.Format | Should Be "General"
                }
            }
            $col = $ws.Cells["C2:C"] # PriorityClass
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [System.Diagnostics.ProcessPriorityClass] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
            $col = $ws.Cells["D2:D"] # TotalProcessorTime
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [TimeSpan] | Should Be $true
            }
            $col | % {
                if ($_.Value -ne $null) {
                    $_.Style.NumberFormat.Format | Should Be "hh:mm:ss"
                }
                else {
                    $_.Style.NumberFormat.Format | Should Be "General"
                }
            }
        }
        $xlPkg.Save()
        $xlPkg.Dispose()

        $columnOptions = @{
            "*" = @{ ForceText = $true }
        }
        $xlPkg = $process | Export-Excel $workbook -PassThru -ColumnOptions $columnOptions
        It "Responds to -ColumnOptions" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $col = $ws.Cells["A:D"]
            $col | Select-Object | % {
                $_.Style.NumberFormat.Format | Should Be "General"
            }
        }
        $xlPkg.Save()
        $xlPkg.Dispose()
        # Invoke-Item $workbook
    }

    Context "With Export-Excel and [valuetype]" {
        $workbook = New-TestWorkbook
        $xlPkg = "12 January 1984" | Export-Excel $workbook -PassThru
        It "Produces [datetime] for date/time [string]" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $col = $ws.Cells["A1"] # First cell.
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [DateTime] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be (Get-DateFormatDefault)
            }
        }
        $xlPkg.Save()
        $xlPkg.Dispose()
        # Invoke-Item $workbook

        $xlPkg = "0123" | Export-Excel $workbook -PassThru
        It "Produces [string] for numeric [string] with leading zeroes" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $col = $ws.Cells["A1"] # First cell.
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [string] | Should Be $true
                $_ | Should Be "0123"
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
        }
        $xlPkg.Save()
        $xlPkg.Dispose()
        # Invoke-Item $workbook

        $xlPkg = "123" | Export-Excel $workbook -PassThru
        It "Produces [double] for numeric [string]" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $col = $ws.Cells["A1"] # First cell.
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [double] | Should Be $true
                $_ | Should Be 123
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
        }
        $xlPkg.Save()
        $xlPkg.Dispose()
        # Invoke-Item $workbook

        $xlPkg = "123", 456, "034", $true, (Get-Date), [long]678, "12 January 1984" | Export-Excel $workbook -PassThru
        It "Supports multi-valuetype array (with automatic string conversions)" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            & {
                $col = $ws.Cells["A1"]
                $col.Value -is [double] | Should Be $true # Automatic conversion from string to double.
                $col.Value | Should Be 123
                $col.Style.NumberFormat.Format | Should Be "General"
            }
            & {
                $col = $ws.Cells["A2"]
                $col.Value -is [int] | Should Be $true # No conversion.
                $col.Value | Should Be 456
                $col.Style.NumberFormat.Format | Should Be "General"
            }
            & {
                $col = $ws.Cells["A3"]
                $col.Value -is [string] | Should Be $true # Automatic conversion chose to remain as string.
                $col.Value | Should Be "034"
                $col.Style.NumberFormat.Format | Should Be "General"
            }
            & {
                $col = $ws.Cells["A4"]
                $col.Value -is [bool] | Should Be $true
                $col.Value | Should Be $true # No conversion.
                $col.Style.NumberFormat.Format | Should Be "General"
            }
            & {
                $col = $ws.Cells["A5"]
                $col.Value -is [datetime] | Should Be $true # No conversion.
                $col.Style.NumberFormat.Format | Should Be (Get-DateFormatDefault)
            }
            & {
                $col = $ws.Cells["A6"]
                $col.Value -is [long] | Should Be $true
                $col.Value | Should Be 678 # No conversion.
                $col.Style.NumberFormat.Format | Should Be "General"
            }
            & {
                $col = $ws.Cells["A7"]
                $col.Value -is [datetime] | Should Be $true # Automatic conversion from string to datetime.
                $col.Style.NumberFormat.Format | Should Be (Get-DateFormatDefault)
            }
        }
        $xlPkg.Save()
        $xlPkg.Dispose()
        # Invoke-Item $workbook

        $columnOptions = @{
            1 = @{ IgnoreText = $true }
        }
        $xlPkg = "123", 456, "034", $true, (Get-Date), [long]678, "12 January 1984" | Export-Excel $workbook -ColumnOptions $columnOptions -PassThru
        It "Supports multi-valuetype array (with no string conversions)" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            & {
                $col = $ws.Cells["A1"]
                $col.Value -is [string] | Should Be $true
                $col.Value | Should Be "123" # No automatic conversion of strings.
                $col.Style.NumberFormat.Format | Should Be "General"
            }
            & {
                $col = $ws.Cells["A2"]
                $col.Value -is [int] | Should Be $true
                $col.Value | Should Be 456
                $col.Style.NumberFormat.Format | Should Be "General"
            }
            & {
                $col = $ws.Cells["A3"]
                $col.Value -is [string] | Should Be $true
                $col.Value | Should Be "034" # No automatic conversion of strings.
                $col.Style.NumberFormat.Format | Should Be "General"
            }
            & {
                $col = $ws.Cells["A4"]
                $col.Value -is [bool] | Should Be $true
                $col.Value | Should Be $true
                $col.Style.NumberFormat.Format | Should Be "General"
            }
            & {
                $col = $ws.Cells["A5"]
                $col.Value -is [datetime] | Should Be $true
                $col.Style.NumberFormat.Format | Should Be (Get-DateFormatDefault)
            }
            & {
                $col = $ws.Cells["A6"]
                $col.Value -is [long] | Should Be $true
                $col.Value | Should Be 678
                $col.Style.NumberFormat.Format | Should Be "General"
            }
            & {
                $col = $ws.Cells["A7"]
                $col.Value -is [string] | Should Be $true
                $col.Value | Should Be "12 January 1984" # No automatic conversion of strings.
                $col.Style.NumberFormat.Format | Should Be "General"
            }
        }
        $xlPkg.Save()
        $xlPkg.Dispose()
        # Invoke-Item $workbook


        $columnOptions = @{
            1 = @{ IgnoreText = $true }
        }
        $xlPkg = "123", "456", "034" | Export-Excel $workbook -ColumnOptions $columnOptions -PassThru
        It "Produces [string] for numeric [string] with -ColumnOptions" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            & {
                $col = $ws.Cells["A1"]
                $col.Value -is [string] | Should Be $true
                $col.Style.NumberFormat.Format | Should Be "General"
            }
            & {
                $col = $ws.Cells["A2"]
                $col.Value -is [string] | Should Be $true
                $col.Style.NumberFormat.Format | Should Be "General"
            }
            & {
                $col = $ws.Cells["A3"]
                $col.Value -is [string] | Should Be $true
                $col.Style.NumberFormat.Format | Should Be "General"
            }
        }
        $xlPkg.Save()
        $xlPkg.Dispose()
        # Invoke-Item $workbook

        $xlPkg = ([long]123) | Export-Excel $workbook -PassThru
        It "Produces [long] for [long] input" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $col = $ws.Cells["A1"] # First cell.
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [long] | Should Be $true
                $_ | Should Be 123
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
        }
        $xlPkg.Save()
        $xlPkg.Dispose()
        # Invoke-Item $workbook

        $xlPkg = 123 | Export-Excel $workbook -PassThru
        It "Produces [int] for [int] input" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $col = $ws.Cells["A1"] # First cell.
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [int] | Should Be $true
                $_ | Should Be 123
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
        }
        $xlPkg.Save()
        $xlPkg.Dispose()
        # Invoke-Item $workbook
    }

    Context "With Export-Excel and CSV data" {
        $workbook = New-TestWorkbook
        $csvData = @"
Name, ID, Age, Birthday
Aa, 123, 82, 12 January 1984
BB, 012, 34, 12 August 1955
CC, 901, 44, 30 May 1901
"@ | ConvertFrom-Csv
        $xlPkg = $csvData | Export-Excel $workbook -DateTimeFormat "mmm/dd/yyyy" -PassThru
        It "Produces Excel data with correct formatting" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $col = $ws.Cells["A2:A"] # Name
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [string] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
            $col = $ws.Cells["B1"] # ID
            $col | Select-Object -ExpandProperty Value | % {
                $_ | Should Be "ID"
                $_ -is [string] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
            $col = $ws.Cells["B2"]
            $col | Select-Object -ExpandProperty Value | % {
                $_ | Should Be 123
                $_ -is [double] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
            $col = $ws.Cells["B3"]
            $col | Select-Object -ExpandProperty Value | % {
                $_ | Should Be "012"
                $_ -is [string] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
            $col = $ws.Cells["B4"]
            $col | Select-Object -ExpandProperty Value | % {
                $_ | Should Be 901
                $_ -is [double] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
            $col = $ws.Cells["C2:C"] # Age
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [double] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
            $col = $ws.Cells["D2:D"] # Birthday
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [DateTime] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "mmm/dd/yyyy"
            }
        }
        $xlPkg.Save()
        $xlPkg.Dispose()
        # Invoke-Item $workbook

        $columnOptions = @{
            ID = @{ IgnoreText = $true}
            3 = @{ IgnoreText = $true }
            Birthday = @{ DateTimeFormat = "mmm/dd/yyyy" }
        }

        $xlPkg = $csvData | Export-Excel $workbook -ColumnOptions $columnOptions -PassThru
        It "Produces Excel data with -ColumnOptions" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $col = $ws.Cells["B2:B"] # ID
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [string] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
            $col = $ws.Cells["C2:C"] # Age
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [string] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
            $col = $ws.Cells["D2:D"] # Birthday
            $col | % {
                $_.Style.NumberFormat.Format | Should Be "mmm/dd/yyyy"
            }
        }
        $xlPkg.Save()
        $xlPkg.Dispose()
        # Invoke-Item $workbook

        $columnOptions = @{
            "*" = @{ IgnoreText = $true }
        }

        $xlPkg = $csvData | Export-Excel $workbook -ColumnOptions $columnOptions -PassThru
        It "Produces Excel data with -ColumnOptions *" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $col = $ws.Cells["A2:A"] # Name
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [string] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
            $col = $ws.Cells["B2:B"] # ID
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [string] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
            $col = $ws.Cells["C2:C"] # Age
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [string] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
            $col = $ws.Cells["D2:D"] # Birthday
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [string] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
        }
        $xlPkg.Save()
        $xlPkg.Dispose()
        # Invoke-Item $workbook
    }

    Remove-TestWorkbook
}

