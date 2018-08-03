function Get-ObjectTitles {
    [CmdletBinding()]
    param(
        $Object
    )
    $ArrayList = New-Object System.Collections.ArrayList
    Write-Verbose "Get-ObjectTitles - ObjectType $($Object.GetType())"
    foreach ($Title in $Object.PSObject.Properties) {
        Write-Verbose "Get-ObjectTitles - Value added to array: $($Title.Name)"
        $ArrayList.Add($Title.Name) | Out-Null
    }
    Write-Verbose "Get-ObjectTitles - Array size: $($ArrayList.Count)"
    return $ArrayList
}
function Get-ObjectData {
    [CmdletBinding()]
    param(
        $Object,
        $Title,
        [switch] $DoNotAddTitles
    )
    $ArrayList = New-Object System.Collections.ArrayList
    $Values = $Object.$Title
    Write-Verbose "Get-ObjectData1: Title $Title Values: $Values"
    if ((Get-ObjectCount $values) -eq 1 -and $DoNotAddTitles -eq $false) {
        $ArrayList.Add("$Title - $Values") | Out-Null
    } else {
        if ($DoNotAddTitles -eq $false) { $ArrayList.Add($Title) | Out-Null }
        foreach ($Value in $Values) {
            $ArrayList.Add("$Value") | Out-Null
        }
    }
    Write-Verbose "Get-ObjectData2: Title $Title Values: $(Get-ObjectCount $ArrayList)"
    return $ArrayList
}
function Get-ObjectCount {
    [CmdletBinding()]
    param(
        $Object
    )
    return $($Object | Measure-Object).Count
}

function Get-ObjectTypeInside {
    [CmdletBinding()]
    param(
        $Object
    )
    if ($Object -ne $null) {
        $ObjectType = $Object.GetType().Name
        if ($ObjectType -eq 'Object[]') {
            if ((Get-ObjectCount $Object) -gt 0) {
                $ObjectTypeInsider = $Object[0].GetType().Name

            }
        } else {
            $ObjectTypeInsider = $ObjectType
        }
    }

    return $ObjectTypeInsider
}
function Get-ObjectType {
    [CmdletBinding()]
    param(
        $Object,
        [string] $ObjectName
    )
    $Return = [ordered] @{}
    $Return.ObjectName = $ObjectName
    if ($Object -ne $null) {
        $ObjectType = $Object.GetType().Name
        $ObjectTypeBaseName = $Object.GetType().BaseType
        $Return.ObjectTypeName = $ObjectType
        $Return.ObjectTypeBaseName = $ObjectTypeBaseName
        #$Return.ObjectTypeIsNested = $Object.GetType().IsNested

        #if ($ObjectType -eq 'Object[]') {
        if ((Get-ObjectCount $Object) -gt 0) {
            $Return.ObjectTypeInsiderName = if ($Object[0] -ne $null) { $Object[0].GetType().Name } else { '' }
            $Return.ObjectTypeInsiderBaseName = if ($Object[0] -ne $null) { $Object[0].GetType().BaseType } else { '' }
            #     $Return.ObjectTypeInsiderIsNested = if ($Object[0] -ne $null) { $Object[0].GetType().IsNested } else { '' }
        } else {
            $Return.ObjectTypeInsiderName = ''
            $Return.ObjectTypeInsiderBaseName = ''
            #     $Return.ObjectTypeInsiderIsNested = ''
        }
    } else {
        $Return.ObjectTypeName = ''
        $Return.ObjectTypeBaseName = ''
        # $Return.ObjectTypeIsNested = ''
        $Return.ObjectTypeInsiderName = ''
        $Return.ObjectTypeInsiderBaseName = ''
        #     $Return.ObjectTypeInsiderIsNested = ''
    }
    return  $Return.ForEach( {[PSCustomObject]$_})
}