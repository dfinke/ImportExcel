Set-StrictMode -Version Latest

$script = Join-Path $PSScriptRoot "$(Split-Path -Leaf $PSCommandPath)".Replace(".Tests.ps1", ".ps1")

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

        It "Leaves numeric strings as text when using -AsText switch" {
            "12345" | & $script -AsText | % {
                $_.Value -is [string] | Should Be $true
                $_.Value | Should Be "12345"
                $_.Format | Should Be "General"
            }
        }

        It "Converts date strings to the default date format" {
            $date = Get-Date
            "$date" | & $script| % {
                $_.Value -is [datetime] | Should Be $true
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

    Context "Piping [datetime] inputs" {
        It "Outputs the same [datetime]" {
            $date = Get-Date
            $date | & $script | % {
                $_.Value -is [datetime] | Should Be $true
                $_.Value | Should Be $date
                $_.Format | Should Be "m/d/yy h:mm"
            }
        }
    }

    Context "Piping [double] inputs" {
        It "Outputs the same [double]" {
            123 | & $script | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 123
                $_.Format | Should Be "General"
            }
        }
    }

    Context "Piping other numeric inputs" {
        It "Outputs [long] as [double]" {
            ([long]123) | & $script | % {
                $_.Value -is [double] | Should Be $true
                $_.Value | Should Be 123
                $_.Format | Should Be "General"
            }
        }
    }

    Context "Piping other input types" {
        It "Outputs [bool] as [string]" {
            $true | & $script | % {
                $_.Value -is [string] | Should Be $true
                $_.Value | Should Be "True"
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
            $csvData | Select-Object -ExpandProperty ID | & $script -AsText | % {
                $_.Value -is [string] | Should Be $true
                $_.Format | Should Be "General"
            }
            $csvData | Select-Object -ExpandProperty Age | & $script | % {
                $_.Value -is [double] | Should Be $true
                $_.Format | Should Be "General"
            }
            $csvData | Select-Object -ExpandProperty Birthday | & $script | % {
                $_.Value -is [datetime] | Should Be $true
                $_.Format | Should Be "m/d/yy h:mm"
            }
        }
    }

    Context "Piping Get-Process data" {
        $process = Get-Process powershell
        It "Converts property values to appropriate types" {
            $process | Select-Object -ExpandProperty StartTime | & $script | % {
                $_.Value -is [datetime] | Should Be $true
                $_.Format | Should Be "m/d/yy h:mm"
            }
            $process | Select-Object -ExpandProperty Id | & $script | % {
                $_.Value -is [double] | Should Be $true
                $_.Format | Should Be "General"
            }
            $process | Select-Object -ExpandProperty ProcessName | & $script | % {
                $_.Value -is [string] | Should Be $true
                $_.Format | Should Be "General"
            }
            $process | Select-Object -ExpandProperty Handles | & $script | % {
                $_.Value -is [double] | Should Be $true
                $_.Format | Should Be "General"
            }
        }
        It "Can interpret numbers as strings" {
            $process | Select-Object -ExpandProperty Id | & $script | % {
                $_.Value -is [double] | Should Be $true
                $_.Format | Should Be "General"
            }
            $process | Select-Object -ExpandProperty Id | & $script -AsText | % {
                $_.Value -is [string] | Should Be $true
                $_.Format | Should Be "General"
            }
        }
    }

}
