$PSVersionTable.PSVersion

if ((Get-Module -ListAvailable pester) -eq $null) {
    Install-Module -Name Pester -Repository PSGallery -Force
}

Invoke-Pester -Script $PSScriptRoot\__tests__