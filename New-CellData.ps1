<#
.SYNOPSIS

Create a new column options table.

#>
function New-ColumnOptions {
    [CmdletBinding()]
    param(
        [hashtable]$Table=@{},
        [string]$NumberFormat="General",
        [string]$DateTimeFormat="mmm/dd/yyyy hh:mm:ss"
    )
    [pscustomobject]@{
        Cache = [ordered]@{}
        Options = $Table
        Default = [ordered]@{
            IgnoreText = $false
            ForceText = $false
            DateTimeFormat = $DateTimeFormat
            NumberFormat = $NumberFormat
        }
    }
}
<#
.SYNOPSIS

Returns the column options for a particular column.

#>
function Get-ColumnOptions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [pscustomobject]$Options,
        [string]$ColumnIndex,
        [string]$ColumnName
    )

    $opts = [pscustomobject]($Options.Cache[$ColumnIndex])
    if ($opts -eq $null) {
        $opts = [pscustomobject]($Options.Default)
        $optsTable = $Options.Options
        if ($optsTable -ne $null) {
            foreach ($optPair in $optsTable.GetEnumerator()) {
                $optKey = $optPair.Name
                $optVal = $optPair.Value
                if ($ColumnIndex -like $optKey) {
                    foreach ($optValPair in $optVal.GetEnumerator()) {
                        # "by index $optKey $optValPair" | Out-Host
                        $opts."$($optValPair.Name)" = $optValPair.Value
                    }
                }
                if ($ColumnName -like $optKey) {
                    foreach ($optValPair in $optVal.GetEnumerator()) {
                        # "by name $optKey $optValPair" | Out-Host
                        $opts."$($optValPair.Name)" = $optValPair.Value
                    }
                }
            }
            $Options.Cache[$ColumnIndex] = $opts
        }
    }
    else {
        # "Found $opts" | Out-Host
    }
    $opts
}

<#
.SYNOPSIS

This function determines the desired cell value and format for the incoming
data. Incoming strings can be interpreted as [DateTime], [double], percentages
etc, unless the -IgnoreText switch is used. Incoming objects that are not
strings, can have their formatting detected to some degree, or they can be
converted to string without further interpretation using the -ForceText
switch. This function found inspiration from the LoadFrom and ConvertData
methods of EPPlus/ExcelRangeBase.cs.

.PARAMETER ForceText

This means that incoming objects (string or non-string) will be treated as
text, and their string values will not be interpreted. The cell value will be
a string and the format will be "General" Useful when piping non-string
objects, like [DateTime] etc.

.PARAMETER IgnoreText

This means that incoming string objects will not be converted to their
interpreted type and they won't be formatted. The cell value will be a string
and the format will be "General". Useful when you don't like the interpreted
results or have to keep strings as strings.

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
        [switch]$ForceText,
        [switch]$IgnoreText
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
            if ($ForceText.IsPresent) {
                makeOut "" "General"
            }
            else {
                makeOut $null "General"
            }
        }
        else {
            foreach ($itemObject in $Objects) {

                if ($ForceText.IsPresent) {
                    $itemObject = $itemObject.ToString()
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
                elseif ($itemObject -is [string] -and !$ForceText.IsPresent -and !$IgnoreText.IsPresent) {
                    $out = fromString $itemObject
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
