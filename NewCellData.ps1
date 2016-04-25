<#
.SYNOPSIS

This function receives arbitrary objects and returns the types and formats
expected by EPPlus. It tries not to loose too much information about the
incoming type.

.DESCRIPTION

This function receives arbitrary objects and returns the types and formats
expected by EPPlus. It found inspiration in the LoadFrom and ConvertData
methods of EPPlus/ExcelRangeBase.cs. The EPPlus type and format architecture
is based on [string], [double] and [datetime] objects, further described by
number formats ("General", etc.).

.PARAMETER SkipText

Do not interpret strings.

.EXAMPLE

PS> "123" | .\NewCellData.ps1 | Select-Object -ExpandProperty Value

Returns [double] 123.

.EXAMPLE

PS> "0123" | .\NewCellData.ps1 | Select-Object -ExpandProperty Value

Returns [string] 0123.

.EXAMPLE

PS> "1/1/13 1:10" | .\NewCellData.ps1 | Select-Object -ExpandProperty Value

Returns [datetime]. 

.EXAMPLE

PS> Get-Date | .\NewCellData.ps1

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
[CmdletBinding()]
param(
    [Parameter(ValueFromPipeline=$true)]
    [object[]]$Objects,
    [string]$NumberFormat="General",
    [System.Globalization.NumberStyles]$NumberStyles=[System.Globalization.NumberStyles]::Any,
    # [string]$DateTimeFormat=[System.Globalization.DateTimeFormatInfo]::CurrentInfo.ShortDatePattern,
    [string]$DateTimeFormat="mmm/dd/yyyy hh:mm:ss",
    # [string]$DateTimeFormat="mm/dd/yy hh:mm",
    # [string]$DateTimeFormat="m/d/yy h:mm", # Doesn't work as expected.
    # [string]$DateTimeFormat="yyyy/mm/dd hh:mm",
    # [string]$DateTimeFormat="dd-mmm-yy",
    [string]$TimeSpanFormat="hh:mm:ss",
    [System.Globalization.DateTimeStyles]$DateTimeStyles=[System.Globalization.DateTimeStyles]::None,
    [string]$PercentageFormat="0.00##\%",
    [switch]$SkipText
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
        $double = 0
        if ($Value -match "^\s+.*") {
            # The string starts with whitespace.
        }
        elseif ($Value -match ".*\s+$") {
            # The string ends with whitespace.
        }
        elseif ($Value.StartsWith("0")) {
            # The string starts with zero. It could be a [DateTime] string.
        }
        elseif ([double]::TryParse($Value, $NumberStyles, [System.Globalization.NumberFormatInfo]::InvariantInfo, [ref]$double)) {
            makeOut $double $NumberFormat
        }
    }
    # A helper that determines the formatting of a string.
    function fromString([string]$Value) {
        $out = asDate($Value)
        if ($out -eq $null) {
            $out = asPercentage($Value)
            if ($out -eq $null) {
                $out = asDouble($Value)
                if ($out -eq $null) {
                    $out = makeOut $Value $NumberFormat
                }
            }
        }
        $out
    }
}
process {
    Set-StrictMode -Version Latest
    $itemList = $_
    $itemList | % {
        $itemObject = $_
        $out = makeOut $itemObject "General"

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
            else {
                # Accept the default.
            }
        }
        elseif ($itemObject -is [string]) {
            if (!$SkipText.IsPresent) {
                $out = fromString $itemObject
            }
            else {
                # Accept the default.
            }
        }
        else {
            # Accept the default.
        }

        $out
    }
}
end {
}
