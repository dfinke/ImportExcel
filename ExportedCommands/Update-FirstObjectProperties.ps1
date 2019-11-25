[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='Update*', Justification='Does not change system state')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Scope='Function', Target='Update*', Justification='Property would be incorrect')]

param()

function Update-FirstObjectProperties {
    [CmdletBinding()]
    param()
    try {
        $Union = @()
        $Input | ForEach-Object {
            If ($Union.Count) {
                $_ | Get-Member | Where-Object {-not ($Union[0] | Get-Member $_.Name)} | ForEach-Object {
                    $Union[0] | Add-Member -MemberType NoteProperty -Name $_.Name -Value $Null
                }
            }
            $Union += $_
        }
        $Union
    }
    catch {
        throw "Failed updating the properties of the first object: $_"
    }
}