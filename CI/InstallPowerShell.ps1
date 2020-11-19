<#
    .SYNOPSIS
    Installs PowerShell Core on Windows.
#>
[CmdLetBinding()]
Param
(
    # Version to install in the format from the .msi, for example "7.0.0-preview.1"
    [String]$Version
)
$ErrorActionPreference = 'Stop'

if (-not $Version) {
    $Version = (Invoke-RestMethod https://raw.githubusercontent.com/PowerShell/PowerShell/master/tools/metadata.json).StableReleaseTag
}
$Version = $Version -replace "^v",""

'[Progress] Downloading PowerShell Core.'
$MsiPath = Join-Path $env:TEMP "PowerShell-$Version-win-x64.msi"
[System.Net.WebClient]::new().DownloadFile("https://github.com/PowerShell/PowerShell/releases/download/v$Version/PowerShell-$Version-win-x64.msi", $MsiPath)

'[Progress] Installing PowerShell Core.'
Start-Process 'msiexec.exe' -Wait -ArgumentList "/i $MsiPath /quiet"
Remove-Item -Path $MsiPath
$PowerShellFolder = $Version[0]
if ($Version -like "*preview*") {
    $PowerShellFolder += '-preview'
}
$env:Path = "$env:ProgramFiles\PowerShell\$PowerShellFolder;$env:Path"
'[Progress] PowerShell Core Installed.'