#Requires -Module Pester
#Requires -Module ImportExcel
Set-StrictMode -Version Latest

. $PSScriptRoot\Export-Excel.ps1

$script = Join-Path $PSScriptRoot "$(Split-Path -Leaf $PSCommandPath)".Replace(".Tests.ps1", ".ps1")

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

Describe "NewCellData" {

    Context "Piping [string] inputs" {

        It "Converts numeric strings to [double]" {
            "12345" | & $script | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 12345
                $_.Format | Should Be "General"
            }
        }

        It "Leaves numeric strings with leading zeroes as strings" {
            "012345" | & $script | % {
                $_.Value -is [string] | Should Be $true
                $_.Value | Should Be "012345"
                $_.Format | Should Be "General"
            }
        }

        It "Leaves numeric strings as text when using -SkipText switch" {
            "12345" | & $script -SkipText | % {
                $_.Value -is [string] | Should Be $true
                $_.Value | Should Be "12345"
                $_.Format | Should Be "General"
            }
        }

        It "Converts date strings to the default date format" {
            $date = Get-Date
            "$date" | & $script| % {
                $_.Value -is [DateTime] | Should Be $true
                "$($_.Value)" | Should Be "$date"
                $_.Format | Should Be (Get-DateFormatDefault)
            }
        }

        It "Leaves numeric strings with starting and trailing whitespace as strings" {
            " 12345" | & $script | % {
                $_.Value -is [string] | Should Be $true
                $_.Value | Should Be " 12345"
                $_.Format | Should Be "General"
            }
            "12345 " | & $script | % {
                $_.Value -is [string] | Should Be $true
                $_.Value | Should Be "12345 "
                $_.Format | Should Be "General"
            }
        }

        It "Converts percentage strings to numbers" {
            "90%" | & $script | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 0.9
                $_.Format | Should Not Be "General"
            }
        }
    }

    Context "Piping [DateTime] inputs" {
        It "Outputs the same [DateTime]" {
            $date = Get-Date
            $date | & $script | % {
                $_.Value -is [DateTime] | Should Be $true
                $_.Value | Should Be $date
                $_.Format | Should Be (Get-DateFormatDefault)
            }
        }
    }

    Context "Piping numeric value type inputs" {
        It "Outputs [int] for [int] input" {
            123 | & $script | % {
                $_.Value -is [int] | Should Be $true
                $_.Value | Should Be 123
                $_.Format | Should Be "General"
            }
        }
        It "Outputs [double] for [double] input" {
            ([double]123) | & $script | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 123
                $_.Format | Should Be "General"
            }
        }
        It "Outputs [long] for [long] input" {
            ([long]123) | & $script | % {
                $_.Value -is [long] | Should Be $true
                $_.Value | Should Be 123
                $_.Format | Should Be "General"
            }
        }
    }

    Context "Piping other value type inputs" {
        It "Outputs [bool] for [bool] input" {
            $true | & $script | % {
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
            $csvData | Select-Object -ExpandProperty Name | & $script | % {
                $_.Value -is [string] | Should Be $true
                $_.Format | Should Be "General"
            }
            $csvData | Select-Object -ExpandProperty ID | & $script -SkipText | % {
                $_.Value -is [string] | Should Be $true
                $_.Format | Should Be "General"
            }
            $csvData | Select-Object -ExpandProperty Age | & $script | % {
                $_.Value -is [double] | Should Be $true
                $_.Format | Should Be "General"
            }
            $csvData | Select-Object -ExpandProperty Birthday | & $script | % {
                $_.Value -is [DateTime] | Should Be $true
                $_.Format | Should Be (Get-DateFormatDefault)
            }
            $csvData | Select-Object -ExpandProperty Birthday | & $script -SkipText | % {
                $_.Value -is [string] | Should Be $true
                $_.Format | Should Be (Get-DateFormatDefault)
            }
        }
    }

    Context "Piping Get-Process data" {
        $process = Get-Process powershell
        It "Converts property values to appropriate types" {
            $process | Select-Object -ExpandProperty StartTime | & $script | % {
                $_.Value -is [DateTime] | Should Be $true
                $_.Format | Should Be (Get-DateFormatDefault)
            }
            $process | Select-Object -ExpandProperty Id | & $script | % {
                $_.Value -is [int] | Should Be $true
                $_.Format | Should Be "General"
            }
            $process | Select-Object -ExpandProperty ProcessName | & $script | % {
                $_.Value -is [string] | Should Be $true
                $_.Format | Should Be "General"
            }
            $process | Select-Object -ExpandProperty Handles | & $script | % {
                $_.Value -is [int] | Should Be $true
                $_.Format | Should Be "General"
            }
        }
    }

    Context "With Export-Excel and PsCustomObject" {
        $workbook = New-TestWorkbook
        $process = Get-Process | Select-Object Id, StartTime, PriorityClass, TotalProcessorTime
        $xlPkg = $process | Export-Excel $workbook -PassThru
        It "Produces correctly formatted sheet for Get-Process" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $col = $ws.Cells["A2:A"] # Id
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [int] | Should Be $true
            }
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
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

        $xlPkg = "123","456","034" | Export-Excel $workbook -TextColumnList 1 -PassThru
        It "Produces [string] for numeric [string] with -TextColumnList" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $col = $ws.Cells["A:A"] # First column.
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

        $xlPkg = $csvData | Export-Excel $workbook -TextColumnList ID, 3 -DateTimeFormat "mmm/dd/yyyy" -PassThru
        It "Produces Excel data with -TextColumnList" {
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
        }
        $xlPkg.Save()
        $xlPkg.Dispose()
        # Invoke-Item $workbook
    }

    Remove-TestWorkbook
}
