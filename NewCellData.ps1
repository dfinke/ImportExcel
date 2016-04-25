<#
.SYNOPSIS

This function receives arbitrary objects and returns the types and formats
expected by EPPlus. This function always outputs a [string], [double] or
[datetime], with a calculated or user-defined number format.

.DESCRIPTION

This function receives arbitrary objects and returns the types and formats
expected by EPPlus. It found inspiration in the LoadFrom and ConvertData
methods of EPPlus/ExcelRangeBase.cs. The EPPlus type and format architecture
is based on [string], [double] and [datetime] objects, further described by
number formats ("General", etc.). This function then always outputs a
[string], [double] or [datetime], and tries to determine the number format. 

If the input value is already a numeric type, it returns it as a [double]. If
it is already a [datetime], it returns it as such, but tries to determine the
date format. If it is not a number or a [datetime], it casts the object to a
[string], then tries to interpret the value as a [double] or [datetime]. If
all conversion attempts fail, it returns a [string] with number format
"General".

If the $AsText switch is provided, then input objects are treated as [string]
and returned as [string]. The objects may still be interpreted to determine
the formatting. Sometimes input numbers are actually identification strings
that should be treated as strings and not as numbers.

.PARAMETER AsText

Treats the input data as [string] and returns it without interpretation.

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
    [string]$DateTimeFormat="m/d/yy h:mm",
    [System.Globalization.DateTimeStyles]$DateTimeStyles=[System.Globalization.DateTimeStyles]::None,
    [string]$PercentageFormat="0.00##\%",
    [switch]$AsText
)
begin {
    Set-StrictMode -Version Latest
    # A helper function that creates the output object.
    function makeOut($Value, [string]$Format) {
        [PSCustomObject][ordered]@{ Value = $Value; Format = $Format; }
    }
    # A helper that checks for a numeric value type.
    function isNumber($Value) {
        $Value -is [byte]  -or $Value -is [int16]  -or $Value -is [int32]  -or $Value -is [int64]  `
           -or $Value -is [sbyte] -or $Value -is [uint16] -or $Value -is [uint32] -or $Value -is [uint64] `
           -or $Value -is [float] -or $Value -is [double] -or $Value -is [decimal]
    }
}
process {
    Set-StrictMode -Version Latest
    $itemList = $_
    $itemList | % {
        $itemObject = $_
        $out = $null

        if ($AsText.IsPresent) {
            # The user explicitly requested text formatting.
            $out = makeOut "$itemObject" $NumberFormat
        }
        elseif (isNumber $itemObject) {
            $out = makeOut ([double]$itemObject) $NumberFormat
        }
        elseif ($itemObject -is [datetime]) {
            $out = makeOut $itemObject $DateTimeFormat
        }
        else {
            $itemString = "$itemObject" # Ensure that $itemObject is a string, relying on [object]'s ToString method.
            $out = makeOut $itemString $NumberFormat # The default output is a string.
            if ($itemString -match "^\s+.*") {
                # If the string starts with whitespace, then treat it as a string, even if the rest of the characters are numbers.
            }
            elseif ($itemString -match ".*\s+$") {
                # If the string ends with whitespace, then treat it as a string, even if the rest of the characters are numbers.
            }
            elseif ($itemString.StartsWith("0")) {
                # If the string starts with a zero, then don't treat it as a number. But maybe it can still be a [datetime].
                # https://msdn.microsoft.com/en-us/library/9h21f14e(v=vs.110).aspx
                $dateTime = 0
                if ([DateTime]::TryParse($itemString, [System.Globalization.DateTimeFormatInfo]::InvariantInfo, $DateTimeStyles, [ref]$dateTime)) {
                    $out = makeOut $dateTime $DateTimeFormat
                }
            }
            else {
                # Start interpreting [double] and [datetime] values.

                # Try to interpret the string as a [datetime].
                $outDateTime = & {
                    $dateTime = 0
                    if ([DateTime]::TryParse($itemString, [System.Globalization.DateTimeFormatInfo]::InvariantInfo, $DateTimeStyles, [ref]$dateTime)) {
                        makeOut $dateTime $DateTimeFormat
                    }
                }
                if ($outDateTime -ne $null) {
                    # Succeeded as [datetime].
                    $out = $outDateTime
                }
                else {
                    # Not a [datetime]. Try a [double].

                    # Try to iterpret the string as [double] percentage.
                    $outDouble = & {
                        $asPercentage = $false
                        $itemNumber = $itemString
                        if ($itemNumber -match "\s*%\s*$") {
                            # $itemNumber = $itemNumber.Substring(0, $itemNumber.Length - 1)
                            $itemNumber = $itemNumber -replace "\s*%\s*$",""
                            $asPercentage = $true
                        }

                        $double = 0
                        if ([double]::TryParse($itemNumber, $NumberStyles, [System.Globalization.NumberFormatInfo]::InvariantInfo, [ref]$double)) {
                            if ($asPercentage) {
                                makeOut ($double/100.0) $PercentageFormat
                            }
                            else {
                                makeOut $double $NumberFormat
                            }
                        }
                    }
                    if ($outDouble -ne $null) {
                        # Succeeded as [double].
                        $out = $outDouble
                    }
                }
            }
        }
        $out
    }
}
end {
}
