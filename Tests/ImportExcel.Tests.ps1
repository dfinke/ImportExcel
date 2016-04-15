#Requires -Module Pester
#Requires -Module ImportExcel
Set-StrictMode -Version Latest

# Import-Module ImportExcel -Force

function New-TestWorkbook {
    $testWorkbook = Join-Path $PSScriptRoot test.xlsx
    if (Test-Path $testWorkbook) {
        rm $testWorkbook -Force
    }
    $testWorkbook
}

function Remove-TestWorkbook {
    Write-Host "Removing test workbook."
    New-TestWorkbook | Out-Null
}

function New-TestDataCsv {
    @"
ID,Product,Quantity,Price,Total
12001,Nails,37,3.99,147.63
12002,Hammer,5,12.10,60.5
12003,Saw,12,15.37,184.44
01200,Drill,20,8,160  
00120,Crowbar,7,23.48,164.36
true,Bla,7,82,12
false,Bla,7,82,12
2009-05-01 14:57:32.8,Yay,1,3,2
"@ | ConvertFrom-Csv 
}

function ConvertTo-Typed {
    <#
        .Synopsis
        Interprets the properties of a PSCustomObject and converts them to
        their closest matching type. Strings and doubles are currently
        supported. Times and dates could be supported at a later stage.
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true)]
        [pscustomobject[]]$Objects,
        [hashtable]$TypeMap
    )
    begin {
    }
    process {
        $object = [ordered]@{}
        $_.PSObject.Properties | % {
            $name = $_.Name
            $value = $_.Value

            # If the property has a type entry in the $TypeMap, then try to force that type.
            $usedTypeMap = $false
            if ($TypeMap -ne $null) {
                $type = $TypeMap.Get_Item($name)
                if ($type -ne "" -and $type -ne $null) {
                    $usedTypeMap = $true
                    if ($type -eq [string]) {
                        $value = "$value"
                    } 
                    elseif ($type -eq [double]) {
                        [double]$double = 0.0
                        if ([double]::TryParse($value, [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::InvariantInfo, [ref]$double)) {
                            $value = $double
                        } else {
                            throw "Double could not be forced. Sorry."
                        }
                    }
                    else {
                        throw "Type not yet supported. Sorry."
                    }
                }
            }

            if (!$usedTypeMap) {

                if ($value.GetType().Name -eq "object") {
                    # The value is of an unspecified type, so let's treat it as a string.
                    # Offer a switch that would disallow conversion to string for general objects.
                    Write-Debug "The value '$value' is an unspecified type, treating it as a string."
                    $value = "$value"
                }

                if ($value -is [string]) {
                    # The value is a string. This is the case for all imported CSV data.
                    # A string can represent anything. Let's see if we can recognize any standard types.
                    if ($value.Trim().StartsWith("0")) {
                        [datetime]$datetime = Get-Date
                        if ([datetime]::TryParse($value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$datetime)) {
                            # 00:00:01 could be the start of a date/time.
                            $value = $datetime
                        } 
                        else {
                            # For now, we keep the value as a raw string.
                        }
                    }
                    else {
                        # The string doesn't start with a zero.
                        # Let's see if the string can be a [double], [bool], [datetime] etc.
                        [double]$double = 0.0
                        [bool]$bool = $false
                        [datetime]$datetime = Get-Date
                        if ([double]::TryParse($value, [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::InvariantInfo, [ref]$double)) {
                            $value = $double
                        }
                        elseif ([bool]::TryParse($value, [ref]$bool)) {
                            $value = $bool
                        }
                        elseif ([datetime]::TryParse($value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$datetime)) {
                            $value = $datetime
                        }
                        else {
                            # For now, we keep the value as a raw string.
                        }
                    }
                }
                elseif ($value -is [pscustomobject[]]) {
                    # Call this recursively. What will happen?
                    $value | ConvertTo-Typed
                }
                elseif ($value -is [array]) {
                    throw "[array] not implemented. Sorry."
                }
                elseif ($value -is [valuetype]) {
                    throw "[valuetype] not implemented. Sorry."
                }
                elseif ($value -is [object]) {
                    throw "[object] not implemented. Sorry."
                }
                else {
                    throw "Unknown type not implemented. Sorry."
                }
            }

            Write-Debug $value.GetType()
            $object.Add($name, $value)
        }
        $psobject = [pscustomobject]$object
        $psobject
    }
    end {
    }
}

Describe "ExportSimple" {
    $workbook = New-TestWorkbook
    $csvData = New-TestDataCsv

    Context "When importing CSV data" {
        It "The type of every property is 'string'" {
            $csvData | % {
                $_.PSObject.Properties | % {
                    $_.Value -is "string" | Should Be $true
                }
            }
        }
    }

    Context "When piping CSV data to ConvertTo-Typed" {
        It "The type of every property is dependent on its content" {
            $excelData = $csvData | ConvertTo-Typed
            $excelData[2] | Select-Object -ExpandProperty ID | Should Be 12003
            ($excelData[2] | Select-Object -ExpandProperty ID) -is [double] | Should Be $true
            $excelData[4] | Select-Object -ExpandProperty ID | Should Be "00120"
            ($excelData[4] | Select-Object -ExpandProperty ID) -is [string] | Should Be $true
            $excelData[5] | Select-Object -ExpandProperty ID | Should Be $true
            ($excelData[5] | Select-Object -ExpandProperty ID) -is [bool] | Should Be $true
            $excelData[6] | Select-Object -ExpandProperty ID | Should Be $false
            ($excelData[6] | Select-Object -ExpandProperty ID) -is [bool] | Should Be $true
            ($excelData[7] | Select-Object -ExpandProperty ID) -is [datetime] | Should Be $true
        }
        It "The type of every property can be forced by property name" {
            # $excelData = $csvData | ConvertTo-Typed
            $excelData = $csvData | ConvertTo-Typed -TypeMap @{ ID=[string]; Quantity=[double] }
            $excelData | Select-Object -ExpandProperty ID | % {
                $_ -is [string] | Should Be $true
            }
            $excelData | Select-Object -ExpandProperty Quantity | % {
                $_ -is [double] | Should Be $true
            }

            # The next section throws an exception because not every ID property could be converted to double
            { $csvData | ConvertTo-Typed -TypeMap @{ ID=[double]; } } | Should Throw
        }
        It "The type of every property can be forced globally" {
            # $excelData = $csvData | ConvertTo-Typed
            $false | Should Be $true
        }
    }

    Context "When piping CSV data to Export-Excel" {
        It "Numeric strings are exported as numbers and not text" {
            $xlPkg = $csvData | Export-Excel $workbook -PassThru
            $ws = $xlPkg.Workbook.WorkSheets[1]

            $csvData[2] | Select-Object -ExpandProperty ID | Should Be "12003"
            $ws.Cells["A4"].Value | Should Be 12003

            $csvData[4] | Select-Object -ExpandProperty ID | Should Be "00120"
            $ws.Cells["A6"].Value | Should Be 120
            $ws.Cells["A6"].Value | Should Not Be "00120"

            $xlPkg.Save()
            $xlPkg.Dispose()
        }
    }

    Remove-TestWorkbook
}
