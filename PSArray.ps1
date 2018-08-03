function New-ArrayList {
    [CmdletBinding()]
    param()
    $List = New-Object System.Collections.ArrayList
    <#
    Mathias Rørbo Jessen:
        The pipeline will attempt to unravel the list on assignment,
        so you'll have to either wrap the empty arraylist in an array,
        like above, or call WriteObject explicitly and tell it not to, like so:
        $PSCmdlet.WriteObject($List,$false)
    #>
    return , $List
}
function Add-ToArray {
    [CmdletBinding()]
    param(
        [System.Collections.ArrayList] $List,
        [Object] $Element
    )
    Write-Verbose "Add-ToArray - Element: $Element"
    $List.Add($Element) > $null
    #return , $List
}
function Remove-FromArray {
    param(
        [System.Collections.ArrayList] $List,
        [Object] $Element,
        [switch] $LastElement
    )
    if ($LastElement) {
        $LastID = $List.Count - 1
        $List.RemoveAt($LastID) > $null
    } else {
        $List.Remove($Element) > $null
    }
}