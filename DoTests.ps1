$PSVersionTable.PSVersion

## Create the zip before the tests run
## Otherwise the EPPlus.dll is in use after the Pester run
Compress-Archive -Path . -DestinationPath .\ImportExcel.zip

if ((Get-Module -ListAvailable pester) -eq $null) {
    Install-Module -Name Pester -Repository PSGallery -Force
}

$result = Invoke-Pester -Script $PSScriptRoot\__tests__ -Verbose -PassThru

if ($result.FailedCount -gt 0) {
    throw "$($result.FailedCount) tests failed."
}