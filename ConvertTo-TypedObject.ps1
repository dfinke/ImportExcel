<#
    .SYNOPSIS

    Interprets the properties of a PSCustomObject and converts them to their
    closest matching type. The supported types at this stage are: [string],
    [double] and [datetime].

    .PARAMETER TypeMap

    A hashtable associating property names with their desired types, that
    will be applied to the incoming objects.

    .PARAMETER GlobalType

    The desired type for every property of the incoming objects.

    .EXAMPLE

    PS> $csvData | ConvertTo-TypedObject -TypeMap @{ ID=[string]; Quantity=[double] }

    This will ensure that all the ID property values will be of type [string],
    and all Quantity property values will be of type double.

    .EXAMPLE

    PS> $csvData | ConvertTo-TypedObject -GlobalType ([string])

    This will ensure that every property value will be of type [string].
    
#>
function ConvertTo-TypedObject {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true)]
        [pscustomobject[]]$Objects,
        [hashtable]$TypeMap,
        [type]$GlobalType
    )
    begin {
    }
    process {
        $object = [ordered]@{}
        $_.PSObject.Properties | % {
            $name = $_.Name
            $value = $_.Value

            [type]$typeRequested = $null

            # Use the global type if we have one.
            if ($GlobalType -ne $null) {
                $typeRequested = $GlobalType
            }

            # Use the property type if we have one.
            if ($TypeMap -ne $null) {
                $type = $TypeMap.Get_Item($name)
                if ($type -ne "" -and $type -ne $null) {
                    $typeRequested = [type]$type
                }
            }

            # Detect the type.

            if ($typeRequested -eq $null) {

                if ($value.GetType().Name -eq "object") {
                    # The value is of an unspecified type, so let's treat it as a string.
                    # Offer a switch that would disallow conversion to string for general objects.
                    Write-Verbose "The value '$value' is an unspecified type, treating it as a string."
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
                    $value | ConvertTo-TypedObject
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
            else {
                if ($typeRequested -eq [string]) {
                    $value = "$value"
                } 
                elseif ($typeRequested -eq [double]) {
                    [double]$double = 0.0
                    if ([double]::TryParse($value, [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::InvariantInfo, [ref]$double)) {
                        $value = $double
                    } else {
                        throw "Double could not be forced. Sorry."
                    }
                }
                else {
                    throw "Type '$($typeRequested)' not yet supported. Sorry."
                }
            }

            Write-Debug "Adding value '$value' of type '$($value.GetType())'..."
            $object.Add($name, $value)
        }
        $psobject = [pscustomobject]$object
        $psobject
    }
    end {
    }
}
