<#
.SYNOPSIS

This script interprets the cell format and type (as supported by EPPlus) for
the incoming data. It found inspiration from the LoadFrom and ConvertData
methods of EPPlus/ExcelRangeBase.cs.

.PARAMETER ForceText

This will treat incoming data as text, and won't interpret the string values.
If this switch is not specified, then string inputs are tested to see if they
match predefined patterns for dates, percentages and numbers.

.EXAMPLE

PS> "123" | New-CellData | Select-Object -ExpandProperty Value

Returns [double] 123.

.EXAMPLE

PS> "0123" | New-CellData | Select-Object -ExpandProperty Value

Returns [string] 0123.

.EXAMPLE

PS> "123" | New-CellData -ForceText | Select-Object -ExpandProperty Value

Returns [string] "123".

.EXAMPLE

PS> "1/1/13 1:10" | New-CellData | Select-Object -ExpandProperty Value

Returns [datetime]. 

.EXAMPLE

PS> Get-Date | New-CellData

Returns [datetime].

.LINK

https://epplus.codeplex.com/wikipage?title=FAQ
https://epplus.codeplex.com/SourceControl/latest#EPPlus/ExcelRangeBase.cs
https://epplus.codeplex.com/SourceControl/latest#EPPlus/ExcelTextFormat.cs
https://epplus.codeplex.com/SourceControl/latest#EPPlus/Style/ExcelNumberFormat.cs
http://stackoverflow.com/questions/29473920/how-to-set-cell-data-type
http://stackoverflow.com/questions/24933947/excel-date-column-returning-int-using-epplus
http://stackoverflow.com/questions/28591763/epplus-how-to-know-the-format-of-the-worksheet-cell
http://stackoverflow.com/questions/23816043/epplus-date-cell-datatype-not-working
http://stackoverflow.com/questions/9859610/how-to-set-column-type-when-using-epplus
http://www.bartsp34ks.nl/news/powershell-how-to-create-a-net-datatable-and-save-it-to-xml-file/
https://github.com/RamblingCookieMonster/PowerShell/blob/master/Out-DataTable.ps1
https://www.ablebits.com/office-addins-blog/2015/03/26/excel-convert-text-date/
https://support.office.com/en-us/article/Format-numbers-as-dates-or-times-418bd3fe-0577-47c8-8caa-b4d30c528309
http://stackoverflow.com/questions/10928030/in-powershell-how-can-i-test-if-a-variable-holds-a-numeric-value
http://stackoverflow.com/questions/16806849/number-format-in-excel-showing-value-without-multiplying-with-100
https://msdn.microsoft.com/en-us/library/system.globalization.datetimeformatinfo.invariantinfo(v=vs.110).aspx
https://msdn.microsoft.com/en-us/library/91hfhz89(v=vs.110).aspx

#>
function New-CellData {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true)]
        [object[]]$Objects,
        [System.Globalization.NumberStyles]$NumberStyles=[System.Globalization.NumberStyles]::Any,
        [System.Globalization.DateTimeStyles]$DateTimeStyles=[System.Globalization.DateTimeStyles]::None,
        [string]$NumberFormat="General",
        [string]$DateTimeFormat="mmm/dd/yyyy hh:mm:ss",
        [string]$TimeSpanFormat="hh:mm:ss",
        [string]$PercentageFormat="0.00##\%",
        [switch]$ForceText
    )
    begin {
        Set-StrictMode -Version Latest
        # A helper function that creates the output object.
        function makeOut([object]$Value, [string]$Format) {
            [PSCustomObject][ordered]@{ Value = $Value; Format = $Format; }
        }
        # A helper that checks for a numeric value type.
        function isNumber($Value) {
            $Value -is [byte]  -or $Value -is [int16]  -or $Value -is [int32]  -or $Value -is [int64]  `
               -or $Value -is [sbyte] -or $Value -is [uint16] -or $Value -is [uint32] -or $Value -is [uint64] `
               -or $Value -is [float] -or $Value -is [double] -or $Value -is [decimal]
        }
        # A helper that detects if a string value represents a percentage.
        function asPercentage([string]$Value) {
            if ($Value -match "\s*%\s*$") {
                $Value = $Value -replace "\s*%\s*$",""
                $double = 0
                if ([double]::TryParse($Value, $NumberStyles, [System.Globalization.NumberFormatInfo]::InvariantInfo, [ref]$double)) {
                    makeOut ($double/100.0) $PercentageFormat
                }
            }
        }
        # A helper that detects if a string value represents a date.
        function asDate([string]$Value) {
            $dateTime = 0
            if ([DateTime]::TryParse($Value, [System.Globalization.DateTimeFormatInfo]::InvariantInfo, $DateTimeStyles, [ref]$dateTime)) {
                # https://msdn.microsoft.com/en-us/library/system.datetime.tooadate.aspx
                makeOut $dateTime $DateTimeFormat
            }
        }
        # A helper that detects if a string value represents a double.
        function asDouble([string]$Value) {
            if ($Value -match "^[1-9].*[0-9]$") {
                # The string starts with 1-9 and ends with 0-9. This looks like a number.
                $double = 0
                if ([double]::TryParse($Value, $NumberStyles, [System.Globalization.NumberFormatInfo]::InvariantInfo, [ref]$double)) {
                    makeOut $double $NumberFormat
                }
            }
        }
        # A helper that determines the formatting of a string.
        function fromString([string]$Value) {
            $out = asDate($Value)
            if ($out -eq $null) {
                $out = asPercentage($Value)
                if ($out -eq $null) {
                    $out = asDouble($Value)
                }
            }
            $out
        }
    }
    process {
        Set-StrictMode -Version Latest

        if (($Objects -eq $null) -or ($Objects.Count -eq 0)) {
            makeOut $null "General"
        }
        else {
            foreach ($itemObject in $Objects) {
                if ($ForceText.IsPresent) {
                    $itemObject = "$itemObject" # Relies on the built-in ToString method of the object.
                }

                $out = $null
                if ($itemObject -is [valuetype]) {
                    if ($itemObject -is [DateTime]) {
                        # https://msdn.microsoft.com/en-us/library/system.datetime.tooadate.aspx
                        $out = makeOut $itemObject $DateTimeFormat
                    }
                    elseif ($itemObject -is [TimeSpan]) {
                        $out = makeOut $itemObject $TimeSpanFormat
                    }
                    elseif (isNumber $itemObject) {
                        $out = makeOut $itemObject $NumberFormat
                    }
                }
                elseif ($itemObject -is [string]) {
                    if (!$ForceText.IsPresent) {
                        $out = fromString $itemObject
                    }
                }

                if ($out -eq $null) {
                    $out = makeOut $itemObject "General"
                }

                $out
            }
        }
    }
    end {
    }
}
