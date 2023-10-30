function Get-Range  {
    [CmdletBinding()]
    param($Start=0,$Stop,$Step=1)
    for ($idx = $Start; $idx -lt $Stop; $idx+=$Step) {$idx}
}