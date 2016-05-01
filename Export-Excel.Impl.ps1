Set-StrictMode -Version Latest

<#
.SYNOPSIS

A filter that selects a property as a new name.

#>
filter Select-TargetData([string]$Property) {
    $prop = $_ | Select-Object -ExpandProperty $Property
    [PSCustomObject]@{ TargetData = $prop }
}

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

A helper function that parses a double and returns some other helpful information.

.LINK

http://stackoverflow.com/questions/25211148/detect-iformatprovider-or-cultureinfo-from-string-using-vb-net

#>
function New-NumberFromText {
    [CmdletBinding()]
    param(
        [string]$Text,
        # https://msdn.microsoft.com/en-us/library/system.globalization.numberstyles(v=vs.110).aspx
        [System.Globalization.NumberStyles]$NumberStyles=[System.Globalization.NumberStyles]::Any,
        # https://msdn.microsoft.com/en-us/library/system.globalization.numberformatinfo(v=vs.110).aspx
        [System.Globalization.NumberFormatInfo]$NumberFormatInfo=[System.Globalization.NumberFormatInfo]::InvariantInfo,
        [switch]$Offer,
        [string[]]$OfferCultures=@("en-US", "en-GB")
    )

    $result = [PSCustomObject]@{
        Text = $Text
        Number = 0
        ParseOkay = $false
        NumberStyles = $NumberStyles
        NumberToString = $null
        GoodMatch = $false
    }

    $double = 0
    if ([double]::TryParse($result.Text, $result.NumberStyles, $NumberFormatInfo, [ref]$double)) {
        $result.Number = $double
        $result.ParseOkay = $true
        $result.NumberToString = $result.Number.ToString($NumberFormatInfo)
        if ($result.Text -eq $result.NumberToString) {
            # Cheaper test for well-formed $Text values.
            $result.GoodMatch = $true
        }
        else {

            # More expensive test for trickier cases.

            # The $NumberToString value will typically have more info than the
            # $Text value, due to limitations in the binary representation of
            # numbers. In some cases a $Text value of "0,3" produces a
            # $NumberToString value of "3", which is not a good match for the
            # $Text value. Guard against this. The $Text value might also
            # include other symbols, like +, - or currency, that will not be
            # included in NumberToString. Ideally, the incoming $Text value
            # should already be cleaned up as much as possible.

            if ($result.NumberToString -like "*$($result.Text)*") {
                $result.GoodMatch = $true
            }
        }
    }

    # Offer recommendations.

    if ($result.ParseOkay -and !$result.GoodMatch -and $Offer.IsPresent) {

        Write-Warning "Information loss detected. Got number '$($result.Number)' with format '$($result.NumberToString)' for input '$Text'. Use -Verbose for NumberFormatInfo recommendations."

        if ($VerbosePreference -ne "SilentlyContinue") {
            # Check invariant culture.
            if ($NumberFormatInfo -ne ([System.Globalization.NumberFormatInfo]::InvariantInfo)) {
                $nft = New-NumberFromText -Text $Text -NumberStyles $NumberStyles -NumberFormatInfo ([System.Globalization.NumberFormatInfo]::InvariantInfo)
                Write-Verbose "For number '$($nft.Number)' with format '$($nft.NumberToString)' use NumberFormatInfo 'InvariantInfo'."
            }

            # Check current culture.
            if ($NumberFormatInfo -ne ([System.Globalization.NumberFormatInfo]::CurrentInfo)) {
                $nft = New-NumberFromText -Text $Text -NumberStyles $NumberStyles -NumberFormatInfo ([System.Globalization.NumberFormatInfo]::CurrentInfo)
                Write-Verbose "For number '$($nft.Number)' with format '$($nft.NumberToString)' use NumberFormatInfo 'CurrentInfo'."
            }

            # Check user provided cultures.
            $OfferCultures | % {
                $culture = $_
                $nfi = [CultureInfo]::GetCultureInfo($culture).NumberFormat
                $nft = New-NumberFromText -Text $Text -NumberStyles $NumberStyles -NumberFormatInfo $nfi
                Write-Verbose "For number '$($nft.Number)' with format '$($nft.NumberToString)' use NumberFormatInfo '[CultureInfo]::GetCultureInfo('$culture').NumberFormat'."
            }
        }
    }

    # Write-Verbose $result
    $result
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

Prototype - The prototype for cache entries. It contains the default values.

#>
function New-ColumnOptionsCache {
    [CmdletBinding()]
    param(
        [hashtable]$Table=@{},
        # Number
        [string]$NumberFormat="General",
        # https://msdn.microsoft.com/en-us/library/system.globalization.numberstyles(v=vs.110).aspx
        [System.Globalization.NumberStyles]$NumberStyles=[System.Globalization.NumberStyles]::Any,
        # https://msdn.microsoft.com/en-us/library/system.globalization.numberformatinfo(v=vs.110).aspx
        [System.Globalization.NumberFormatInfo]$NumberFormatInfo=[System.Globalization.NumberFormatInfo]::InvariantInfo,
        # DateTime
        [string]$DateTimeFormat="mmm/dd/yyyy hh:mm:ss",
        # https://msdn.microsoft.com/en-us/library/system.globalization.datetimestyles(v=vs.110).aspx
        [System.Globalization.DateTimeStyles]$DateTimeStyles=[System.Globalization.DateTimeStyles]::None,
        # https://msdn.microsoft.com/en-us/library/system.globalization.datetimeformatinfo(v=vs.110).aspx
        [System.Globalization.DateTimeFormatInfo]$DateTimeFormatInfo=[System.Globalization.DateTimeFormatInfo]::InvariantInfo,
        # TimeSpan
        [string]$TimeSpanFormat="hh:mm:ss"
    )
    [PSCustomObject]@{
        Cache = [ordered]@{}
        Options = $Table
        Prototype = [ordered]@{
            IgnoreText = $false
            ForceText = $false
            DateTimeFormat = $DateTimeFormat
            NumberFormat = $NumberFormat
            NumberStyles = $NumberStyles
            DateTimeStyles = $DateTimeStyles
            NumberFormatInfo = $NumberFormatInfo
            DateTimeFormatInfo = $DateTimeFormatInfo
            TimeSpanFormat = $TimeSpanFormat
            TargetData = $null
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
        [PSCustomObject]$CacheObject,
        [Parameter(Mandatory=$true, HelpMessage="The index of the column whose options we are requesting.")]
        [string]$ColumnIndex,
        [string]$ColumnName,
        [object]$TargetData
    )

    $colOpts = [PSCustomObject]($CacheObject.Cache[$ColumnIndex])
    if ($colOpts -eq $null) {
        $colOpts = [PSCustomObject]($CacheObject.Prototype)
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
                        $optName = $patternOptsPair.Name
                        $colOpts.$optName = $patternOptsPair.Value
                    }
                }
                elseif ($ColumnName -like $colPattern) {
                    # An option entry was found for the $ColumnName.
                    $colPatternOptsTable = $colPatternOptsPair.Value
                    foreach ($patternOptsPair in $colPatternOptsTable.GetEnumerator()) {
                        $optName = $patternOptsPair.Name
                        $colOpts.$optName = $patternOptsPair.Value
                    }
                }
            }
        }

        $CacheObject.Cache[$ColumnIndex] = $colOpts

        Write-Verbose "Column options for '$ColumnIndex'/'$ColumnName': $colOpts)"
    }
    $newOpts = [PSCustomObject]$colOpts
    $newOpts.TargetData = $TargetData
    $newOpts
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
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [object[]]$TargetData,
        # https://msdn.microsoft.com/en-us/library/system.globalization.numberstyles(v=vs.110).aspx
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [System.Globalization.NumberStyles]$NumberStyles=[System.Globalization.NumberStyles]::Any,
        # https://msdn.microsoft.com/en-us/library/system.globalization.datetimestyles(v=vs.110).aspx
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [System.Globalization.DateTimeStyles]$DateTimeStyles=[System.Globalization.DateTimeStyles]::None,
        # https://msdn.microsoft.com/en-us/library/system.globalization.numberformatinfo(v=vs.110).aspx
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [System.Globalization.NumberFormatInfo]$NumberFormatInfo=[System.Globalization.NumberFormatInfo]::InvariantInfo,
        # https://msdn.microsoft.com/en-us/library/system.globalization.datetimeformatinfo(v=vs.110).aspx
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [System.Globalization.DateTimeFormatInfo]$DateTimeFormatInfo=[System.Globalization.DateTimeFormatInfo]::InvariantInfo,
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$NumberFormat="General",
        # https://support.office.com/en-us/article/Format-a-date-the-way-you-want-8e10019e-d5d8-47a1-ba95-db95123d273e?ui=en-US&rs=en-US&ad=US&fromAR=1
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$DateTimeFormat="mmm/dd/yyyy hh:mm:ss",
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$TimeSpanFormat="hh:mm:ss",
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [switch]$ForceText,
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [switch]$IgnoreText
    )
    begin {
        Set-StrictMode -Version Latest
    }
    process {
        Set-StrictMode -Version Latest

        if (($TargetData -eq $null) -or ($TargetData.Count -eq 0)) {
            if ($ForceText.IsPresent) {
                New-ValueFormatPair -Value "" -Format "General"
            }
            else {
                New-ValueFormatPair -Value $null -Format "General"
            }
        }
        else {
            foreach ($itemObject in $TargetData) {
                # Write-Verbose "TargetData is $itemObject"

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
                            $decSep = $NumberFormatInfo.NumberDecimalSeparator
                            if ($itemObject -notmatch "^[0][^\$decSep]+|^[\s]+|[\s]+$") {

                                # "001" is not a valid number, but "0.01" is.
                                # "123" can be interpeted as a number, but the
                                # values " 123"  or "123 " are intentional
                                # strings due to the whitespace. Try to
                                # support these cases.

                                # We also want to support different cultures
                                # (i.e. 0.1 vs 0,1 or 1,200 vs 1.200). This
                                # can already be explicitly set in the
                                # parameters. It makes sense that users of the
                                # library will be working with data formatted
                                # in their CurrentCulture. Ideally, when
                                # converting an initial string to double, the
                                # reverse conversion from double to string
                                # should lead to the same initial string. If
                                # it doesn't, issue a warning, so that the
                                # user can know that they should reformat
                                # their data or use a different
                                # $NumberFormatInfo value.

                                # We also want to support + and -, and perhaps
                                # even currency symbols. See $NumberStyles
                                # documentation for the built-in method of
                                # handling these situations. We might need to
                                # extract the $NumberFormat from the string in
                                # these cases.

                                # If the conversion produces a value that is
                                # not reversible, then don't prefer the
                                # conversion, but retain the original
                                # unconverted value.

                                # Offer recommendations so that the user can
                                # try more optimal $NumberFormatInfo
                                # configurations.

                                $nftUser = New-NumberFromText -Text $itemObject -NumberStyles $NumberStyles -NumberFormatInfo $NumberFormatInfo -Offer
                                if ($nftUser.ParseOkay) {
                                    New-ValueFormatPair -Value $nftUser.Number -Format $NumberFormat
                                }
                            }
                            else {
                                Write-Warning "Leading zeros or whitespace detected. Value '$itemObject' will not be treated as a number."
                            }
                        }
                    }

                    # Is $itemObject a DateTime?
                    if ($out -eq $null) {
                        $out = & {
                            $dateTime = 0
                            if ([DateTime]::TryParse($itemObject, $DateTimeFormatInfo, $DateTimeStyles, [ref]$dateTime)) {
                                # https://msdn.microsoft.com/en-us/library/system.datetime.tooadate.aspx
                                New-ValueFormatPair -Value $dateTime -Format $DateTimeFormat
                            }
                        }
                    }

                    # TODO: Is $itemObject a TimeSpan?
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
