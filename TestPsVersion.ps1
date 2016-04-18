<#
.SYNOPSIS

Tests for a particular PowerShell version.

.PARAMETER Version

The version to test for.

.EXAMPLE

PS> . .\Test-PsVersion.ps1 4

This will return true if the PowerShell version is 4 or greater.

#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [double]$Version
)
$psver = $PSVersionTable.PSVersion
$check = $null
if ($psver.Major -ne $null) {
    $check = $psver.Major
}
else {
    $check = $psver
}
([double]"$check") -ge $Version
