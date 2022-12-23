If ($Null -eq (Get-Module -ListAvailable PlatyPS)) {
    Install-Module -Name PlatyPS -Force -Confirm:$False
}
Import-Module platyPS
Get-ChildItem $PSScriptRoot -Directory | ForEach-Object {
    New-ExternalHelp -Path $_.FullName  -OutputPath (Join-Path $PSScriptRoot "..\$($_.Name)") -Force -Verbose
}