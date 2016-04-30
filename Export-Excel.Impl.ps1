<#

.SYNOPSIS

A helper function that creates a cell value format pair.

#>
function New-ValueFormatPair {
    [CmdletBinding()]
    param(
        [object]$Value,
        [string]$Format
    )
    [PSCustomObject][ordered]@{ Value = $Value; Format = $Format; }
}

<#
.SYNOPSIS

Tests to see if an object is a numeric type.

#>
function Test-NumericType {
    [CmdletBinding()]
    param(
        [object]$Object
    )
    if ($Object -is [ValueType]) {
        $Object -is [double] -or $Object -is [int16]  -or $Object -is [int32]  -or $Object -is [int64]  `
           -or $Object -is [sbyte] -or $Object -is [uint16] -or $Object -is [uint32] -or $Object -is [uint64] `
           -or $Object -is [float] -or $Object -is [byte] -or $Object -is [decimal]
    }
    else {
        $false
    }
}

<#
.SYNOPSIS

Create a new column options cache object.

.DESCRIPTION

The properties of the returned object are:

Cache - Stores the overrides for each column. Calculated once.

Options - A reference to the incoming table. Used to create the cache.

Prototype - The prototype for cache entries:

    IgnoreText - Determines if column text will be converted to appropriate
    objects, or ignored.

    ForceText - Determines whether column data should be treated as strings.

    DateTimeFormat - The format for DateTime objects.

    NumberFormat - The format for numeric data types.

#>
function New-ColumnOptionsCache {
    [CmdletBinding()]
    param(
        [hashtable]$Table=@{},
        [string]$NumberFormat="General",
        [string]$DateTimeFormat="mmm/dd/yyyy hh:mm:ss"
    )
    [pscustomobject]@{
        Cache = [ordered]@{}
        Options = $Table
        Prototype = [ordered]@{
            IgnoreText = $false
            ForceText = $false
            DateTimeFormat = $DateTimeFormat
            NumberFormat = $NumberFormat
        }
    }
}
<#
.SYNOPSIS

Returns the column options for a particular column, given a cache created by
New-ColumnOptionsCache. Column options are strongly associated with the column
index.

.PARAMETER CacheObject

A cache object created by New-ColumnOptionsCache.

.PARAMETER ColumnIndex

The index of the column whose options we are requesting.

.PARAMETER ColumnName

The optional name of the property associated with the column.

#>
function Get-ColumnOptions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, HelpMessage="A cache object created by New-ColumnOptionsCache.")]
        [pscustomobject]$CacheObject,
        [Parameter(Mandatory=$true, HelpMessage="The index of the column whose options we are requesting.")]
        [string]$ColumnIndex,
        [string]$ColumnName
    )

    $colOpts = [pscustomobject]($CacheObject.Cache[$ColumnIndex])
    if ($colOpts -eq $null) {
        $colOpts = [pscustomobject]($CacheObject.Prototype)
        if ($CacheObject.Options -ne $null) {
            # The cache contains no options associated for the particular column
            # index. Let's create the options for it based on the cache prototype.
            $colOptsTable = $CacheObject.Options
            foreach ($colPatternOptsPair in $colOptsTable.GetEnumerator()) {
                # Iterate through every option pair in the options table, and find
                # all options associated with the given $ColumnIndex and
                # $ColumnName. Use those options to update our prototype. If an
                # option exists in the option table that does not exist in the
                # prototype, then throw an error, but make sure that handle
                # -ErrorAction SilentlyContinue won't break anything.
                $colPattern = $colPatternOptsPair.Name
                if ($ColumnIndex -like $colPattern) {
                    # An option entry was found for the $ColumnIndex.
                    $colPatternOptsTable = $colPatternOptsPair.Value
                    foreach ($patternOptsPair in $colPatternOptsTable.GetEnumerator()) {
                        # $colOpts is a PsCustomObject, not a table, so
                        # $assignments to non-existing properties will throw an
                        # $exception.
                        $optName = $patternOptsPair.Name
                        $colOpts.$optName = $patternOptsPair.Value
                    }
                }
                elseif ($ColumnName -like $colPattern) {
                    # An option entry was found for the $ColumnName.
                    $colPatternOptsTable = $colPatternOptsPair.Value
                    foreach ($patternOptsPair in $colPatternOptsTable.GetEnumerator()) {
                        # $colOpts is a PsCustomObject, not a table, so
                        # $assignments to non-existing properties will throw an
                        # $exception.
                        $optName = $patternOptsPair.Name
                        $colOpts.$optName = $patternOptsPair.Value
                    }
                }
            }
        }
        $CacheObject.Cache[$ColumnIndex] = $colOpts
    }
    $colOpts
}

<#
.SYNOPSIS

This function determines the desired cell value and format for the incoming
data. Incoming strings can be interpreted as [DateTime], [double], etc, unless
the -IgnoreText switch is used. Incoming objects that are not strings, can
have their formatting detected to some degree, or they can be converted to
string without further interpretation using the -ForceText switch. This
function found inspiration from the LoadFrom and ConvertData methods of
EPPlus/ExcelRangeBase.cs.

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
        [switch]$ForceText,
        [switch]$IgnoreText
    )
    begin {
        Set-StrictMode -Version Latest
    }
    process {
        Set-StrictMode -Version Latest

        if (($Objects -eq $null) -or ($Objects.Count -eq 0)) {
            if ($ForceText.IsPresent) {
                New-ValueFormatPair -Value "" -Format "General"
            }
            else {
                New-ValueFormatPair -Value $null -Format "General"
            }
        }
        else {
            foreach ($itemObject in $Objects) {

                if ($ForceText.IsPresent) {
                    $itemObject = $itemObject.ToString()
                }

                $out = $null
                if ($itemObject -is [ValueType]) {
                    if ($itemObject -is [DateTime]) {
                        # https://msdn.microsoft.com/en-us/library/system.datetime.tooadate.aspx
                        $out = New-ValueFormatPair -Value $itemObject -Format $DateTimeFormat
                    }
                    elseif ($itemObject -is [TimeSpan]) {
                        $out = New-ValueFormatPair -Value $itemObject -Format $TimeSpanFormat
                    }
                    elseif (Test-NumericType -Object $itemObject) {
                        $out = New-ValueFormatPair -Value $itemObject -Format $NumberFormat
                    }
                }
                elseif ($itemObject -is [string] -and !$ForceText.IsPresent -and !$IgnoreText.IsPresent) {

                    # Is $itemObject a double?
                    if ($out -eq $null) {
                        $out = & {
                            if ($itemObject -notmatch "^[0]+|^[\s]+|[\s]+$") {
                                # The value does not start with a zero, and has no
                                # whitespace at the beginning and end. Let's try to
                                # parse it as a number.
                                $double = 0
                                if ([double]::TryParse($itemObject, $NumberStyles, [System.Globalization.NumberFormatInfo]::InvariantInfo, [ref]$double)) {
                                    New-ValueFormatPair -Value $double -Format $NumberFormat
                                }
                            }
                        }
                    }

                    # Is $itemObject a DateTime?
                    if ($out -eq $null) {
                        $out = & {
                            $dateTime = 0
                            if ([DateTime]::TryParse($itemObject, [System.Globalization.DateTimeFormatInfo]::InvariantInfo, $DateTimeStyles, [ref]$dateTime)) {
                                # https://msdn.microsoft.com/en-us/library/system.datetime.tooadate.aspx
                                New-ValueFormatPair -Value $dateTime -Format $DateTimeFormat
                            }
                        }
                    }
                }

                if ($out -eq $null) {
                    $out = New-ValueFormatPair -Value $itemObject -Format "General"
                }

                $out
            }
        }
    }
    end {
    }
}
