#Requires -Module Pester
#Requires -Module ImportExcel
Set-StrictMode -Version Latest

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
                $_.Format | Should Be "m/d/yy h:mm"
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
                $_.Format | Should Be "m/d/yy h:mm"
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
                $_.Format | Should Be "m/d/yy h:mm"
            }
            $csvData | Select-Object -ExpandProperty Birthday | & $script -SkipText | % {
                $_.Value -is [string] | Should Be $true
                $_.Format | Should Be "m/d/yy h:mm"
            }
        }
    }

    Context "Piping Get-Process data" {
        $process = Get-Process powershell
        It "Converts property values to appropriate types" {
            $process | Select-Object -ExpandProperty StartTime | & $script | % {
                $_.Value -is [DateTime] | Should Be $true
                $_.Format | Should Be "m/d/yy h:mm"
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

    Context "With Export-Excel" {
        $workbook = New-TestWorkbook
        $process = Get-Process | Select-Object Id, StartTime, PriorityClass, TotalProcessorTime
        $xlPkg = $process | Export-Excel $workbook -PassThru
        It "Produces correctly formatted sheet for Get-Process" {
            $ws = $xlPkg.Workbook.WorkSheets[1]
            $col = $ws.Cells["A2:A"] # Id
            $col | Select-Object -ExpandProperty Value | % {
                $_ -is [double] | Should Be $true
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
                    $_.Style.NumberFormat.Format | Should Be "m/d/yy h:mm"
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
            $col | Select-Object -ExpandProperty Style | % {
                $_.NumberFormat.Format | Should Be "General"
            }
        }
        $xlPkg.Save()
        $xlPkg.Dispose()
        # Invoke-Item $workbook
        Remove-TestWorkbook
    }

}
