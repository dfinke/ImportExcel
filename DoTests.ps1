
Get-Module -ListAvailable pester | Out-Host


return
if ((Get-Module -ListAvailable pester) -eq $null) {
    Install-Module -Name Pester -Repository PSGallery -Force
}

Invoke-Pester -Script $PSScriptRoot\UnitTests\